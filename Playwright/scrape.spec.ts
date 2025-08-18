import { test } from '@playwright/test';
import crypto from 'crypto';
import { writeFile } from 'fs/promises';
import path from 'path';
import * as XLSX from 'xlsx';

test.use({ browserName: 'chromium' });
test.setTimeout(120_000); // all tests in this file use 120s

test('scrape all products', async ({ page }) => {
    // Add multiple URLs to this array. All pages will be scraped and the
    // results will be combined into a single output.json / output.xlsx.
    const urls = [
        'https://phannacomputershop.com/cat/desktop-all-in-1/',
        'https://phannacomputershop.com/cat/new-desktop/',
        'https://phannacomputershop.com/cat/new-laptop/',
        'https://phannacomputershop.com/cat/used-laptop/',
        'https://phannacomputershop.com/cat/used-desktop/',
        'https://phannacomputershop.com/cat/pc-part/'
    ];


    // deterministic / unique code helpers (8 chars A-Z0-9)
    const usedCodes = new Set<string>();
    const fingerprintToCode = new Map<string, string>();
    const CHARS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';

    type PropCodeGenerateForProduct = {
        name: string
        category: string | null
        cpu: string | null
        ram: string | null
        storage: string | null
        psu: string | null
        case: string | null
        gpu: string | null
    };

    function generateRandomCode(): string {
        while (true) {
            let out = '';
            for (let i = 0; i < 8; i++) {
                out += CHARS[crypto.randomInt(0, CHARS.length)];
            }
            if (!usedCodes.has(out)) {
                usedCodes.add(out);
                return out;
            }
        }
    }

    function normalizeKeyForFingerprint(prop: PropCodeGenerateForProduct) {
        // include all relevant properties in the fingerprint (normalized)
        const fields: (keyof PropCodeGenerateForProduct)[] = ['name', 'category', 'cpu', 'ram', 'storage', 'psu', 'case'];
        const parts = fields.map(k => {
            const v = (prop[k] ?? '') as string;
            return v.toLowerCase().replace(/\s+/g, ' ').trim();
        });
        return parts.join('||');
    }

    function codeFromHash(hashHex: string): string {
        try {
            const base36 = BigInt('0x' + hashHex).toString(36).toUpperCase().replace(/[^A-Z0-9]/g, '');
            const cand = (base36 + '00000000').slice(0, 8);
            if (!usedCodes.has(cand)) {
                usedCodes.add(cand);
                return cand;
            }
        } catch (e) {
            // fall through to random
        }
        return generateRandomCode();
    }

    function getCodeForProduct(prop: PropCodeGenerateForProduct) {

        const key = normalizeKeyForFingerprint(prop);
        if (fingerprintToCode.has(key)) return fingerprintToCode.get(key)!;
        const hash = crypto.createHash('sha256').update(key).digest('hex');
        const code = codeFromHash(hash);
        fingerprintToCode.set(key, code);
        return code;
    }
    const results: {
        code: string;
        name: string;
        brand: string | null;
        category: string | null;
        price: string;
        image: string | null;
        specs: string[];
        rawHtml: string | null;
    }[] = [];

    const pushedCodes = new Set<string>();

    for (const url of urls) {
        console.log('Processing', url);
        try {
            await page.goto(url, { timeout: 120_000, waitUntil: 'load' });
        } catch (err) {
            console.error('Failed to load', url, err);
            continue; // go to next URL
        }

        // extract the category for the current page (if available)
        const categoryPage = await page.$$('.container-fluid.clearfix .menu-main-menu-container ul li a');
        let category: string | null = null;
        for (const link of categoryPage) {
            const aria = await link.getAttribute('aria-current');
            if (aria === 'page') {
                const span = await link.$('span');
                category = span ? (await span.evaluate(s => s.textContent?.trim())) ?? null : null;

                // normalize to title case (capitalize each word; preserve hyphenated parts)
                if (category) {
                    const toTitleCase = (str: string) =>
                        str
                            .trim()
                            .split(/\s+/)
                            .map(word =>
                                word
                                    .split('-')
                                    .map(part => part ? (part[0].toUpperCase() + part.slice(1).toLowerCase()) : '')
                                    .join('-')
                            )
                            .join(' ');
                    category = toTitleCase(category);
                }
                break;
            }
        }

        const items = await page.$$('.site-content .container');

        for (const item of items) {
            const brandHandle = await item.$('.section-title h2 span');
            const rawBrand = brandHandle ? (await brandHandle.evaluate(el => el.textContent?.trim())) ?? null : null;
            const brand = rawBrand ? rawBrand.split('-')[0].trim() : null;

            const products = await item.$$('.product-item');
            for (const product of products) {
                // title name default
                const nameHandle = await product.$('.title-and-rating h2');
                let description: string[] = [];
                let rawHtml: string | null = null;
                const listHandle = await product.$('.list');
                if (listHandle) {
                    // Capture the full innerHTML (do not trim) so we keep markup and structure intact.
                    rawHtml = (await listHandle.evaluate(el => el.innerHTML ?? null)) ?? null;

                    // Prefer individual logical lines. Some <li> use <br> inside, so split those into separate lines.
                    // This evaluates in the page context to preserve HTML -> text conversion reliably.
                    description = await listHandle.$$eval('li', els =>
                        els.flatMap(li => {
                            // turn <br> tags into newlines, then strip remaining tags and split by newlines
                            const html = li.innerHTML || '';
                            const withBreaks = html.replace(/<br\s*\/?>/gi, '\n');
                            const d = document.createElement('div');
                            d.innerHTML = withBreaks;
                            const text = (d.textContent || '').trim();
                            return text.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
                        })
                    );

                    // Fallback: if no <li> children exist, split the container text by newlines.
                    if (!description || description.length === 0) {
                        const text = (await listHandle.evaluate(el => el.textContent?.trim())) || '';
                        description = text.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
                    }

                    // Normalize each description line:
                    // - replace NBSP with regular space
                    // - collapse multiple whitespace into single spaces
                    // - normalize colons to "Key: Value" (no space before, single after)
                    // - remove spaces directly inside parentheses: "( AIO )" -> "(AIO)"
                    description = description.map(line => {
                        if (!line) return line;
                        let v = line.replace(/\u00A0/g, ' ').trim();
                        v = v.replace(/\s*[:：]\s*/g, ': ');      // normalize colon spacing (handles fullwidth too)
                        v = v.replace(/\s+/g, ' ');              // collapse multiple spaces
                        v = v.replace(/\(\s*(.*?)\s*\)/g, '($1)'); // normalize parentheses spacing
                        return v.trim();
                    });
                }

                const MAX_NAME_LEN = 256;
                const visibleTitle = nameHandle ? (await nameHandle.evaluate(el => el.textContent?.trim())) : null;

                // declare spec vars here so they are available later when generating the code
                let cpuVal: string | null = null;
                let ramVal: string | null = null;
                let gpu: string | null = null;
                let storageVal: string | null = null;
                let psuVal: string | null = null;
                let caseVal: string | null = null;

                // mornitor display
                let resolutionVal: string | null = null;
                // helper to normalize resolution strings
                const normalizeResolution = (s: string) => {
                    return s
                        .replace(/\u00A0/g, ' ')
                        .replace(/●/g, ' ')
                        .replace(/\s+/g, ' ')
                        .replace(/\s*:\s*/g, ': ')
                        .trim();
                };

                // Prefer a description entry that looks like a title (avoid spec lines like "CPU: ...").
                let name: string = '';
                if (description && description.length > 0) {
                    const rawCandidate = description.find(s => !/[:：]/.test(s) && s.length > 3) ?? description[0];

                    // Normalize whitespace and parentheses:
                    // 1) Replace non-breaking spaces with regular spaces.
                    // 2) Collapse any sequence of whitespace to a single space.
                    // 3) Trim leading/trailing spaces.
                    // 4) Remove spaces directly inside parentheses: "( AIO )" -> "(AIO)".

                    // This name fetch only first line of descripton
                    const cleanedCandidate = rawCandidate
                        ? rawCandidate
                            .replace(/\u00A0/g, ' ')
                            .replace(/\s*[:：]\s*/g, ' ')   // remove any colon and surrounding spaces -> single space
                            .replace(/\s+/g, ' ')
                            .trim()
                            .replace(/\(\s*(.*?)\s*\)/g, '($1)')
                        : rawCandidate;

                    // Keys to extract; storageKeys treated together (first one wins)
                    const storageKeys = ['m2', 'm.2', 'ssd', 'storage', 'hdd', 'pci-e'];

                    // normalized GPU keys (no punctuation/space) so we can compare against keyNormalized
                    const gpuKeys = ['gpu', 'graphics', 'graphics amd', 'graphicsamd', 'vga', 'amd', 'amd radeon', 'intel®', 'nvidia', 'nvidia®'];

                    // Scan description lines for "Key : Value" patterns (case-insensitive),
                    // normalize key by removing non-alphanumerics so variants like "M.2", "m.2", "M.2   " are matched.
                    // Use an index loop so we can look ahead (useful for headings like "Memory" / "Processor")
                    for (let i = 0; i < description.length; i++) {
                        const line = description[i];
                        if (!line) continue;
                        const text = line.replace(/\u00A0/g, ' ').trim();

                        // Monitor resolution detection (examples: "Resolution:QHD 2560 x 1440 at 120Hz",
                        // "Resolution : 1920 x 1080 at 120 Hz", or lines that contain "2560 x 1440")
                        if (!resolutionVal) {
                            const t = normalizeResolution(text);
                            // common patterns: QHD/FHD/UHD/4K or explicit dimensions 1920 x 1080, optionally "at 120Hz"
                            const dimRegex = /\b(QHD|FHD|UHD|4K|HD)\b|\b(\d{3,4}\s*[x×]\s*\d{3,4}(?:\s*at\s*\d+\s*Hz)?)\b/i;
                            const keyRes = t.match(/\bResolution\b[:\s\-]*(.+)$/i);
                            const dimMatch = t.match(dimRegex);
                            if (keyRes && keyRes[1]) {
                                resolutionVal = normalizeResolution(keyRes[1]);
                            } else if (dimMatch) {
                                // prefer the explicit dimension capture group if present
                                resolutionVal = (dimMatch[2] || dimMatch[1] || dimMatch[0]).trim();
                            }
                        }

                        // global regex to capture multiple "Key: Value" pairs in the same line
                        const pairRegex = /([^:：]+?)\s*[:：]\s*([^:：]+?)(?=(?:\s+[^:：]+?\s*[:：])|$)/g;
                        let match: RegExpExecArray | null = null;
                        let foundAny = false;

                        // collect values in order of appearance (used as fallback for cpu)
                        const valuesInLine: string[] = [];
                        while ((match = pairRegex.exec(text)) !== null) {
                            foundAny = true;
                            const rawKey = (match[1] || '').trim();
                            const value = (match[2] || '')
                                .replace(/\u00A0/g, ' ')
                                .replace(/\s+/g, ' ')
                                .trim()
                                .replace(/\(\s*(.*?)\s*\)/g, '($1)');
                            valuesInLine.push(value);
                            const keyNormalized = rawKey.replace(/[^a-z0-9]/gi, '').toLowerCase();

                            if (!cpuVal && keyNormalized === 'cpu') cpuVal = value;
                            if (!ramVal && keyNormalized === 'ram') ramVal = value;
                            if (!psuVal && keyNormalized === 'psu') psuVal = value;
                            if (!caseVal && keyNormalized === 'case') caseVal = value;
                            if (!storageVal && storageKeys.includes(keyNormalized)) storageVal = value;

                            // existing freeform detection inside the pair loop (kept as-is)
                            if (!storageVal) {
                                const freeformStorageRegex = /\b(m\.?2|nvme|pci-?e|pcie|ssd|hdd|storage)\b/i;
                                const freeformMatch = text.match(freeformStorageRegex);
                                if (freeformMatch) {
                                    const token = (freeformMatch[1] || '').toLowerCase();
                                    // Ignore m.2 / m2 when it's part of a unit like 'cd/m2' (brightness),
                                    // i.e. when preceded by a slash. Treat only standalone m.2 as storage.
                                    if ((token === 'm.2' || token === 'm2') && /\/\s*m\.?2\b/i.test(text)) {
                                        // skip - likely 'cd/m2' or similar unit, do not treat as storage
                                    } else {
                                        const cap = text.match(/(\d+(?:\.\d+)?\s*(?:TB|GB))/i);
                                        storageVal = (cap ? cap[0] : text).replace(/\s+/g, ' ').trim();
                                    }
                                }
                            }

                            // GPU key handling (mirror storageKeys flow)
                            if (!gpu && gpuKeys.includes(keyNormalized)) {
                                // if the value is only a brand (e.g. "AMD") prefer the next non-empty line as model
                                const brandOnlyRegex = /^(amd|nvidia|intel|ati)$/i;
                                const nextLine = (description[i + 1] || '').trim();
                                if (brandOnlyRegex.test(value) && nextLine) {
                                    gpu = nextLine.replace(/\s+/g, ' ').trim();
                                } else {
                                    gpu = value;
                                }
                            }

                            if (cpuVal && ramVal && storageVal && psuVal && caseVal) break;
                        }

                        // If no "Key: Value" pairs matched (or storage still not found),
                        // try a more robust freeform storage normalizer that captures PCIe version + capacity.
                if (!storageVal) {
                    const normalizeStorageFromText = (t: string) => {
                                const s = t.replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim();
                                // detect primary type
                                const typeMatch = s.match(/\b(m\.?2|m2|nvme|mvme|ssd|hdd|storage)\b/i);
                                // detect PCIe/PCIE + version (e.g. PCIe 5.0 or PCIe4.0)
                                const pcieMatch = s.match(/\bpcie(?:\s*[-]?\s*|\s*)(\d+(?:\.\d+)?)\b/i) || s.match(/\bpci-?e\b/i);
                                // detect NVMe/MVMe token
                                const nvmeMatch = s.match(/\b(nvme|mvme)\b/i);
                                // capacity
                                const capMatch = s.match(/(\d+(?:\.\d+)?\s*(?:TB|GB))/i);

                                const parts: string[] = [];
                                if (typeMatch) {
                                    const t0 = typeMatch[1] ?? typeMatch[0];
                                    if (/m\.?2|m2/i.test(t0)) parts.push('M.2');
                                    else parts.push(t0.toUpperCase());
                                }
                                if (pcieMatch) {
                                    // pcieMatch[1] exists when there's a numeric version
                                    if (pcieMatch[1]) parts.push(`PCIe ${pcieMatch[1]}`);
                                    else parts.push('PCIe');
                                }
                                if (nvmeMatch && !parts.some(p => /NVME/i.test(p))) parts.push('NVMe');
                                if (capMatch) {
                                    // normalize capacity spacing to e.g. "2TB"
                                    parts.push(capMatch[1].replace(/\s+/g, '').toUpperCase());
                                }

                                // If we have useful parts return them, otherwise fall back to the cleaned line
                                const out = parts.join(' ').replace(/\s+/g, ' ').trim();
                                return out || s;
                            };

                            // run the normalizer on the whole text line
                            const storageTokenMatch = text.match(/\b(m\.?2|m2|nvme|mvme|pci-?e|pcie|ssd|hdd|storage)\b/i);
                            if (storageTokenMatch) {
                                const token = (storageTokenMatch[1] || '').toLowerCase();
                                // ignore m.2 when it's used as a unit like 'cd/m2'
                                if ((token === 'm.2' || token === 'm2') && /\/\s*m\.?2\b/i.test(text)) {
                                    // do nothing - it's likely a unit (brightness), not storage
                                } else {
                                    storageVal = normalizeStorageFromText(text);
                                }
                            }
                        }

                        // If cpu wasn't explicitly provided but other specs were found,
                        // try smarter fallbacks:
                        // 1) find a CPU-like non-key line (contains GHz or common CPU model words)
                        // 2) otherwise take the second non-key description line (index 1)
                        // 3) otherwise fall back to valuesInLine (prefer index 1)
                        if (!cpuVal && (ramVal || storageVal || psuVal || caseVal)) {
                            const cpuLikeRegex = /\b(i[3579]\b|intel|amd|ryzen|xeon|threadripper|athlon)\b|\d+(\.\d+)?\s*GHz/i;
                            const nonKeyLines = (description || []).filter(l => l && !/[:：]/.test(l));
                            let candidateLine: string | null = null;

                            for (const nl of nonKeyLines) {
                                if (cpuLikeRegex.test(nl)) { candidateLine = nl; break; }
                            }
                            if (!candidateLine && nonKeyLines.length > 1) {
                                // user case: index 1 is often the CPU line (e.g. "Ultra 9 285K 3.2GHz")
                                candidateLine = nonKeyLines[1];
                            }
                            if (!candidateLine && valuesInLine.length > 0) {
                                candidateLine = (valuesInLine[1] ?? valuesInLine[0]) || null;
                            }

                            if (candidateLine) {
                                cpuVal = candidateLine
                                    .replace(/\u00A0/g, ' ')
                                    .replace(/\s+/g, ' ')
                                    .trim()
                                    .replace(/\(\s*(.*?)\s*\)/g, '($1)');
                            }
                        }

                        // Extra heuristics: handle headings and neighbor lines
                        // - "Processor" heading often followed by CPU details on same or next line
                        // - "Memory" heading often followed by DDR line on the next line
                        if (!cpuVal) {
                            // "Processor" as heading or inline descriptor
                            if (/\bprocessor\b/i.test(text)) {
                                // prefer same line details after the word "Processor"
                                // otherwise take the next non-empty line
                                const after = text.replace(/.*processor[:\s\-]*?/i, '').trim();
                                const nextLine = (description[i + 1] || '').trim();
                                const pick = after.length > 0 ? after : nextLine;
                                if (pick && /\d+(\.\d+)?\s*GHz|\bcore\b|\bCPU\b|CPUs?/i.test(pick)) {
                                    cpuVal = pick.replace(/\s+/g, ' ').trim();
                                } else if (pick) {
                                    // fallback: take short summary as cpuVal
                                    cpuVal = pick.replace(/\s+/g, ' ').trim();
                                }
                            } else {
                                // lines that mention "CPUs", "core", "GHz" without "Processor"
                                if (/\bCPUs?\b|\bcore(s)?\b|\d+(\.\d+)?\s*GHz\b/i.test(text)) {
                                    cpuVal = text.replace(/\s+/g, ' ').trim();
                                }
                            }
                        }

                        // Graphics/ GPU heading or freeform GPU model lines
                        if (!gpu) {
                            // heading "Graphics" or "Graphics:" -> prefer details on same line after keyword or next non-empty line
                            if (/\bgraphics\b/i.test(text)) {
                                const after = text.replace(/.*graphics[:\s\-]*?/i, '').trim();
                                const nextLine = (description[i + 1] || '').trim();

                                // tokens indicating a full model line (brand+model)
                                const modelTokenRegex = /\b(FirePro|Radeon|GeForce|GTX|RTX|Quadro|RX|Vega|Titan|TITAN|Ti)\b/i;
                                // common brand-only values
                                const brandOnlyRegex = /^(AMD|NVIDIA|INTEL|ATI)$/i;

                                // prefer a detailed model if available on same line or the next line
                                if (after && modelTokenRegex.test(after)) {
                                    gpu = after.replace(/\s+/g, ' ').trim();
                                } else if (nextLine && modelTokenRegex.test(nextLine)) {
                                    gpu = nextLine.replace(/\s+/g, ' ').trim();
                                } else if (after && !brandOnlyRegex.test(after)) {
                                    // if 'after' is non-brand shorthand (e.g. "FirePro D300 ..."), use it
                                    gpu = after.replace(/\s+/g, ' ').trim();
                                } else if (brandOnlyRegex.test(after) && nextLine && nextLine.length) {
                                    // "Graphics  AMD" + next line with model -> prefer next line
                                    gpu = nextLine.replace(/\s+/g, ' ').trim();
                                } else if (after) {
                                    gpu = after.replace(/\s+/g, ' ').trim();
                                } else if (nextLine) {
                                    gpu = nextLine.replace(/\s+/g, ' ').trim();
                                }
                            } else {
                                // detect common GPU brand/model tokens in freeform lines
                                const gpuTokenRegex = /\b(FirePro|Radeon|GeForce|GTX|RTX|Quadro|RX|Vega|AMD|NVIDIA)\b/i;
                                if (gpuTokenRegex.test(text)) {
                                    // prefer the more specific model line (if this line is just "Graphics  AMD", take the next line)
                                    const nextLine = (description[i + 1] || '').trim();
                                    if (/^\s*(AMD|NVIDIA)\s*$/i.test(text) && nextLine && gpuTokenRegex.test(nextLine)) {
                                        gpu = nextLine.replace(/\s+/g, ' ').trim();
                                    } else {
                                        gpu = text.replace(/\s+/g, ' ').trim();
                                    }
                                }
                            }
                        }

                        // Memory heading + adjacent DDR line handling
                        if (!ramVal) {
                            // If this line is a Memory heading, look at next non-empty line
                            if (/^\s*Memory\s*[:\-]?\s*$/i.test(text)) {
                                const nextLine = (description[i + 1] || '').trim();
                                if (nextLine) {
                                    const m1 = nextLine.match(/\b(DDR[345]X?)\b[\s,;:-]*?(\d+(?:\.\d+)?\s*GB)?/i);
                                    const m2 = nextLine.match(/(\d+(?:\.\d+)?\s*GB)[\s,;:-]*?(DDR[345]X?)/i);
                                    if (m1 || m2) {
                                        const ddr = (m1 && m1[1]) || (m2 && m2[2]) || null;
                                        const cap = (m1 && m1[2]) || (m2 && m2[1]) || null;
                                        if (ddr) {
                                            ramVal = (ddr.toUpperCase() + (cap ? ' ' + cap.replace(/\s+/g, '').toUpperCase() : '')).trim();
                                        } else if (cap) {
                                            ramVal = cap.replace(/\s+/g, '').toUpperCase();
                                        }
                                    } else {
                                        // fallback: if nextLine contains a capacity like "16GB" use it
                                        const cap = nextLine.match(/(\d+(?:\.\d+)?\s*GB)/i);
                                        if (cap) ramVal = cap[1].replace(/\s+/g, '').toUpperCase();
                                    }
                                }
                            } else {
                                // If this line itself contains DDR/capacity info, use existing regexes too
                                const m1 = text.match(/\b(DDR[345]X?)\b[\s,;:-]*?(\d+(?:\.\d+)?\s*GB)?/i);
                                const m2 = text.match(/(\d+(?:\.\d+)?\s*GB)[\s,;:-]*?(DDR[345]X?)/i);
                                if (m1 || m2) {
                                    const ddr = (m1 && m1[1]) || (m2 && m2[2]) || null;
                                    const cap = (m1 && m1[2]) || (m2 && m2[1]) || null;
                                    if (ddr) {
                                        ramVal = (ddr.toUpperCase() + (cap ? ' ' + cap.replace(/\s+/g, '').toUpperCase() : '')).trim();
                                    } else if (cap) {
                                        ramVal = cap[1].replace(/\s+/g, '').toUpperCase();
                                    }
                                }
                            }
                        }

                        // Fallback: single pair match (in case regex didn't find multiple pairs)
                        if (!foundAny) {
                            const m = text.match(/^([^:：]+)\s*[:：]\s*(.+)$/);
                            if (m) {
                                const rawKey = m[1].trim();
                                const value = m[2]
                                    .replace(/\u00A0/g, ' ')
                                    .replace(/\s+/g, ' ')
                                    .trim()
                                    .replace(/\(\s*(.*?)\s*\)/g, '($1)');
                                const keyNormalized = rawKey.replace(/[^a-z0-9]/gi, '').toLowerCase();

                                if (!cpuVal && keyNormalized === 'cpu') cpuVal = value;
                                if (!ramVal && keyNormalized === 'ram') ramVal = value;
                                if (!psuVal && keyNormalized === 'psu') psuVal = value;
                                if (!caseVal && keyNormalized === 'case') caseVal = value;
                                if (!storageVal && storageKeys.includes(keyNormalized)) storageVal = value;
                            }
                        }

                        // Detect DDR memory lines (e.g. "8GB DDR4", "DDR5 16GB", "DDR5X 32GB")
                        if (!ramVal) {
                            // Try both orders: "DDR4 8GB" and "8GB DDR4"
                            const m1 = text.match(/\b(DDR[345]X?)\b[\s,;:-]*?(\d+(?:\.\d+)?\s*GB)?/i);
                            const m2 = text.match(/(\d+(?:\.\d+)?\s*GB)[\s,;:-]*?(DDR[345]X?)/i);
                            let ddr: string | null = null;
                            let cap: string | null = null;
                            if (m1) {
                                ddr = m1[1];
                                cap = m1[2] || null;
                            } else if (m2) {
                                cap = m2[1];
                                ddr = m2[2];
                            }
                            if (ddr) {
                                // Normalize to "DDR4 8GB" or "DDR5X" if capacity missing
                                ramVal = (ddr.toUpperCase() + (cap ? ' ' + cap.replace(/\s+/g, '').toUpperCase() : '')).trim();
                            }
                        }

                        if (cpuVal && ramVal && storageVal && psuVal && caseVal) break; // got everything we want
                    }

                    // Build final name:
                    // - If none of cpuVal/ramVal/storageVal were found, use cleanedCandidate only.
                    // - If any were found, append the found specs (normalized) to the cleanedCandidate.
                    // treat resolution as a "monitor spec" when no other hardware specs exist
                    const isLikelyMonitor = !cpuVal && !ramVal && !gpu && !storageVal && !psuVal && !caseVal && Boolean(resolutionVal);
                    const hasSpecs = Boolean(cpuVal || ramVal || storageVal || psuVal || caseVal || (isLikelyMonitor && resolutionVal));
                    if (hasSpecs) {
                        const hasPrimarySpecs = Boolean(cpuVal || ramVal || storageVal || gpu);
                        const partsToAppend = (hasPrimarySpecs
                            ? [cpuVal, ramVal, storageVal, gpu]
                            : [psuVal, caseVal]
                        ).filter(Boolean)
                            .map(p => p!.replace(/\s+/g, ' ').trim());

                        // If this is a monitor (no CPU/RAM/Storage) append resolution as the primary spec
                        if (isLikelyMonitor && resolutionVal) {
                            const res = normalizeResolution(resolutionVal);
                            // avoid duplication
                            if (!partsToAppend.map(x => (x || '').toLowerCase()).includes(res.toLowerCase())) {
                                partsToAppend.unshift(res);
                            }
                        }

                        // new: include visible title (from .title-and-rating h2) as a prefix,
                        // but avoid duplicating it if the cleanedCandidate already starts with it.
                        const baseName = (cleanedCandidate ?? '').replace(/\s+/g, ' ').trim();
                        const prefixText = visibleTitle ? visibleTitle.replace(/\s+/g, ' ').trim() : '';
                        let combined = baseName;

                        if (prefixText) {
                            const baseLower = (baseName || '').toLowerCase();
                            const prefixLower = prefixText.toLowerCase();
                            if (baseLower.length === 0 || !baseLower.startsWith(prefixLower)) {
                                combined = prefixText + ' ' + combined;
                            }
                        }

                        // normalize, dedupe and avoid appending parts already present in base/prefix
                        const normalizeSpecPart = (s: string) => {
                            let p = s.replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim();
                            // canonical tokens
                            p = p.replace(/\bmvme\b/ig, 'NVMe');
                            p = p.replace(/\bnvme\b/ig, 'NVMe');
                            p = p.replace(/\bm\.?2\b/ig, 'M.2');
                            p = p.replace(/\bssd\b/ig, 'SSD');
                            // PCIe version normalization: "PCIe5.0", "PCIe 5.0", "PCI-E4" -> "PCIe 5.0" or "PCIe 4"
                            p = p.replace(/\bpcie\s*[-]?\s*(\d+(?:\.\d+)?)\b/ig, (_m, ver) => `PCIe ${ver}`);
                            p = p.replace(/\bpci-?e\b/ig, 'PCIe');
                            return p.replace(/\s+/g, ' ').trim();
                        };

                        const baseLower = (baseName || '').toLowerCase();
                        const prefixLower = prefixText ? prefixText.toLowerCase() : '';
                        const seen = new Set<string>();
                        const filteredParts: string[] = [];
                        for (const raw of partsToAppend) {
                            const p = normalizeSpecPart(raw);
                            const key = p.toLowerCase();
                            if (!p) continue;
                            if (seen.has(key)) continue;               // internal duplicate (e.g. MVMe vs NVMe)
                            if (baseLower.includes(key) || prefixLower.includes(key)) continue; // already in base/prefix
                            seen.add(key);
                            filteredParts.push(p);
                        }
                        if (filteredParts.length) combined = combined + ' ' + filteredParts.join(' ');

                        // collapse repeated adjacent phrases (handles multi-word repeats)
                        const collapseRepeatedPhrases = (s: string, maxWords = 6) => {
                            const words = s.split(/\s+/).filter(Boolean);
                            let i = 0;
                            while (i < words.length) {
                                let removed = false;
                                const maxLen = Math.min(maxWords, Math.floor((words.length - i) / 2));
                                for (let len = maxLen; len >= 1; len--) {
                                    let same = true;
                                    for (let k = 0; k < len; k++) {
                                        if ((words[i + k] || '').toLowerCase() !== (words[i + len + k] || '').toLowerCase()) {
                                            same = false;
                                            break;
                                        }
                                    }
                                    if (same) {
                                        // remove the second occurrence
                                        words.splice(i + len, len);
                                        removed = true;
                                        break;
                                    }
                                }
                                if (!removed) i++;
                            }
                            return words.join(' ');
                        };

                        combined = collapseRepeatedPhrases(combined);
                        name = combined
                            .replace(/\s+/g, ' ')
                            .replace(/\(\s*(.*?)\s*\)/g, '($1)')
                            .trim();
                    } else {
                        name = (cleanedCandidate ?? '').replace(/\s+/g, ' ').trim();
                    }

                } else if (visibleTitle) {
                    name = visibleTitle;
                } else {
                    const headLink = await product.$('.head a');
                    if (headLink) {
                        const alt = (await headLink.evaluate(a => (a.getAttribute('title') || a.textContent || '').trim())) || null;
                        if (alt) name = alt;
                    }
                    if (!name) {
                        name = brand ?? category ?? 'Unnamed Product';
                    }
                }
                if (name.length > MAX_NAME_LEN) name = name.slice(0, MAX_NAME_LEN);

                const priceHandle = await product.$('.price .sale-price');
                const price = priceHandle ? (await priceHandle.evaluate(el => el.textContent?.trim())) ?? '' : '';

                const aTag = await product.$('.head a');
                const image = aTag ? await aTag.getAttribute('href') : null;

                const specs = description;

                if (name) {
                    // generate deterministic 8-char code based on name + price + category
                    const propGenerateCode: PropCodeGenerateForProduct = { name, category, cpu: cpuVal, ram: ramVal, storage: storageVal, psu: psuVal, case: caseVal, gpu };
                    const code = getCodeForProduct(propGenerateCode);

                    const nameWithCode = `${name} - ${code}`;

                    // skip duplicate products by code
                    if (pushedCodes.has(code)) {
                        //console.log('Skipping duplicate product with code', code, 'name:', name)
                    } else {
                        pushedCodes.add(code)
                        results.push({
                            code,
                            name: nameWithCode,
                            brand,
                            category,
                            price,
                            image,
                            specs,
                            rawHtml
                        });
                    }
                }
            }
        }
    }

    const outPath = path.resolve(__dirname, '..', 'output.json');
    try {
        await writeFile(outPath, JSON.stringify(results, null, 2), 'utf8');
        console.log('Saved results to', outPath);
    } catch (err) {
        console.error('Failed to write results:', err);
    }

    try {
        const headers = [
            '_id', 'ID', 'Code', 'Name', 'Price', 'In Stock', 'Image', 'Images', 'Category', 'Item Type',
            'Description', 'Published', 'Price ID 1', 'Barcode 1', 'Price 1', 'Currency 1', 'Unit Name 1',
            'Unit Rank 1', 'Unit Size 1', 'Type 1', 'Price ID 2', 'Barcode 2', 'Price 2', 'Currency 2',
            'Unit Name 2', 'Unit Rank 2', 'Unit Size 2', 'Type 2'
        ];

        const rows = results.map((r, idx) => {
            const priceNumber = r.price ? String(r.price).replace(/[^0-9.]/g, '') : '';
            const descriptionHtml = (r.rawHtml && r.rawHtml.length)
                ? r.rawHtml
                : (r.specs && r.specs.length) ? `<ul>${r.specs.map(s => `<li>${s}</li>`).join('')}</ul>` : '';
            return {
                _id: '',
                ID: idx + 1,
                Code: r.code,
                Name: r.name,
                Price: priceNumber,
                'In Stock': 0,
                Image: r.image ?? '',
                Images: r.image ? JSON.stringify([r.image.toString()]) : '',
                Category: r.category ?? '',
                'Item Type': 'simple',
                Description: descriptionHtml,
                Published: 1,
                'Price ID 1': '',
                'Barcode 1': '',
                'Price 1': '',
                'Currency 1': '',
                'Unit Name 1': '',
                'Unit Rank 1': '',
                'Unit Size 1': '',
                'Type 1': '',
                'Price ID 2': '',
                'Barcode 2': '',
                'Price 2': '',
                'Currency 2': '',
                'Unit Name 2': '',
                'Unit Rank 2': '',
                'Unit Size 2': '',
                'Type 2': ''
            };
        });

        const aoa: any[][] = [
            headers,
            ...rows.map(row => headers.map(h => row[h] ?? ''))
        ];

        const ws = XLSX.utils.aoa_to_sheet(aoa);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

        const outXlsxPath = path.resolve(__dirname, '..', 'output.xlsx');
        XLSX.writeFile(wb, outXlsxPath);
        console.log('Saved excel to', outXlsxPath);
    } catch (err) {
        console.error('Failed to write excel file:', err);
    }
});