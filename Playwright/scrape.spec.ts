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

    function normalizeKeyForFingerprint(name: string, price: string, category: string | null, image: string | null) {
        const n = (name || '').toLowerCase().replace(/\s+/g, ' ').trim();
        const p = (price || '').replace(/[^0-9.]/g, '').trim(); // numeric price normalized
        const c = (category || '').toLowerCase().replace(/\s+/g, ' ').trim();
        const i = (image || '').toLowerCase().replace(/\s+/g, ' ').trim();
        return `${n}||${p}||${c}||${i}`;
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

    function getCodeForProduct(name: string, price: string, category: string | null, image: string | null) {
        const key = normalizeKeyForFingerprint(name, price, category, image);
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
                    const storageKeys = ['m2', 'ssd', 'storage'];

                    let cpuVal: string | null = null;
                    let ramVal: string | null = null;
                    let storageVal: string | null = null;

                    // Scan description lines for "Key : Value" patterns (case-insensitive),
                    // normalize key by removing non-alphanumerics so variants like "M.2", "m.2", "M.2   " are matched.
                    for (const line of description) {
                        if (!line) continue;
                        const text = line.replace(/\u00A0/g, ' ').trim();

                        // global regex to capture multiple "Key: Value" pairs in the same line
                        const pairRegex = /([^:：]+?)\s*[:：]\s*([^:：]+?)(?=(?:\s+[^:：]+?\s*[:：])|$)/g;
                        let match;
                        let foundAny = false;
                        while ((match = pairRegex.exec(text)) !== null) {
                            foundAny = true;
                            const rawKey = (match[1] || '').trim();
                            const value = (match[2] || '')
                                .replace(/\u00A0/g, ' ')
                                .replace(/\s+/g, ' ')
                                .trim()
                                .replace(/\(\s*(.*?)\s*\)/g, '($1)');
                            const keyNormalized = rawKey.replace(/[^a-z0-9]/gi, '').toLowerCase();

                            if (!cpuVal && keyNormalized === 'cpu') cpuVal = value;
                            if (!ramVal && keyNormalized === 'ram') ramVal = value;
                            if (!storageVal && storageKeys.includes(keyNormalized)) storageVal = value;

                            if (cpuVal && ramVal && storageVal) break;
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
                                if (!storageVal && storageKeys.includes(keyNormalized)) storageVal = value;
                            }
                        }
                        if (cpuVal && ramVal && storageVal) break; // got everything we want
                    }

                    // Build final name:
                    // - If none of cpuVal/ramVal/storageVal were found, use cleanedCandidate only.
                    // - If any were found, append the found specs (normalized) to the cleanedCandidate.
                    const hasSpecs = Boolean(cpuVal || ramVal || storageVal);
                    if (hasSpecs) {
                        const partsToAppend = [cpuVal, ramVal, storageVal]
                            .filter(Boolean)
                            .map(p => p!.replace(/\s+/g, ' ').trim());
                        name = (
                            (cleanedCandidate ?? '').replace(/\s+/g, ' ').trim()
                            + ' ' + partsToAppend.join(' ')
                        )
                            .replace(/\s+/g, ' ')
                            .replace(/\(\s*(.*?)\s*\)/g, '($1)')
                            .trim();
                    } else {
                        name = (cleanedCandidate ?? '').replace(/\s+/g, ' ').trim();
                    }

                    /* console.log('rawCandidate=>', rawCandidate);
                    console.log('titleCandidate=>', cleanedCandidate);
                    console.log('extracted=>', { cpuVal, ramVal, storageVal });
                    console.log('description => ', description);
                    console.log('rawHTML=>', rawHtml);
                    console.log('name=>', name); */


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
                    const priceNumber = price ? String(price).replace(/[^0-9.]/g, '') : '';
                    const code = getCodeForProduct(name, priceNumber, category, image);

                    const nameWithCode = `${name} - ${code}`;

                    // skip duplicate products by code
                    if (pushedCodes.has(code)) {
                        console.log('Skipping duplicate product with code', code, 'name:', name)
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