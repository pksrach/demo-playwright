import { test } from '@playwright/test';
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

    const results: {
        name: string;
        brand: string | null;
        category: string | null;
        price: string;
        image: string | null;
        specs: string[];
        rawHtml: string | null;
    }[] = [];

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
                const nameHandle = await product.$('.title-and-rating h2');
                let productSpecs: string[] = [];
                let rawHtml: string | null = null;
                const listHandle = await product.$('.list');
                if (listHandle) {
                    // Capture the full innerHTML (do not trim) so we keep markup and structure intact.
                    rawHtml = (await listHandle.evaluate(el => el.innerHTML ?? null)) ?? null;
                    // Prefer individual <li> items when present to avoid grabbing large combined blocks.
                    productSpecs = await listHandle.$$eval('li', els => els.map(el => el.textContent?.trim()).filter(Boolean));
                    // Fallback: if no <li> children exist, split the container text by newlines.
                    if (!productSpecs || productSpecs.length === 0) {
                        const text = (await listHandle.evaluate(el => el.textContent?.trim())) || '';
                        productSpecs = text.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
                    }
                }

                const MAX_NAME_LEN = 256;
                const visibleTitle = nameHandle ? (await nameHandle.evaluate(el => el.textContent?.trim())) : null;
                // Prefer a productSpecs entry that looks like a title (avoid spec lines like "CPU: ...").
                let name: string = '';
                if (productSpecs && productSpecs.length > 0) {
                    // Choose the first specs entry that does not contain a colon (likely not a "key: value" spec)
                    const titleCandidate = productSpecs.find(s => !/[:：]/.test(s) && (s.length > 3)) ?? productSpecs[0];
                    name = (titleCandidate ?? '').replace(/\u00A0/g, ' ').trim();
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

                const specs = productSpecs;

                if (name) {
                    results.push({
                        name,
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
                Code: '',
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