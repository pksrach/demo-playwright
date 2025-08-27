import { test } from '@playwright/test';
import crypto from 'crypto';
import { writeFile, mkdir } from 'fs/promises';
import path from 'path';
import * as XLSX from 'xlsx';

test.use({ browserName: 'chromium' });
test.setTimeout(150_000);

test('scrape all products (khmer + english)', async ({ page }) => {
    const urls = ['https://nie.edu.kh/'];

    // map lang code -> folder name
    const languages = [
        { code: 'km', folder: 'khmer' },
        { code: 'en', folder: 'english' },
    ];

    const headers = ['Post link', 'Post date', 'Post title', 'Post description', 'Post category'];

    async function scrapeFromPage(pageInstance: any) {
        const localResults: {
            post_date: string;
            post_title: string;
            post_description: string;
            post_link: string;
            post_category?: string;
        }[] = [];

        // ensure all modules' "load more" controls are clicked until no new posts appear
        try {
            const maxClicks = 200; // safety cap across modules
            let clicks = 0;
            // track total article count on page and wait for it to increase
            let prevTotal = await pageInstance.$$eval('article.jeg_post', (els: any) => els.length).catch(() => 0);

            while (clicks < maxClicks) {
                // find all load-more anchors for modules (e.g. #home-news, #home-activities, etc.)
                const buttons = await pageInstance.$$('.jeg_block_loadmore a');
                let didClick = false;

                for (const btn of buttons) {
                    const visible = await btn.isVisible().catch(() => false);
                    if (!visible) continue;

                    // click and wait for the total number of articles to increase
                    await Promise.all([
                        btn.click().catch(() => null),
                        pageInstance.waitForFunction((sel: string, prev: number) => document.querySelectorAll(sel).length > prev, {}, 'article.jeg_post', prevTotal).catch(() => null),
                    ]);

                    const newTotal = await pageInstance.$$eval('article.jeg_post', (els: any) => els.length).catch(() => prevTotal);
                    if (newTotal > prevTotal) {
                        prevTotal = newTotal;
                        didClick = true;
                        break; // restart scanning buttons from the top
                    }

                    // small pause to allow any JS to settle before next button
                    await pageInstance.waitForTimeout(150);
                }

                if (!didClick) break;
                clicks += 1;
                await pageInstance.waitForTimeout(200);
            }
        } catch (e) {
            // ignore and proceed to scrape whatever is present
        }

        // collect all article elements on the page (covers modules like #home-activities)
        const items = await pageInstance.$$('article.jeg_post');
        for (const item of items) {
            let post_category = '';
            let postTitle = '';
            let postDate = '';
            let postDescription = '';
            let postLink = '';

            try {
                post_category = await item.$eval('.jeg_post_category > span > a', (el: any) => (el.textContent || '').trim());
            } catch {
                post_category = '';
            }

            try {
                postTitle = await item.$eval('.jeg_postblock_content .jeg_post_title a', (el: any) => (el.textContent || '').trim());
            } catch {
                postTitle = '';
            }

            try {
                postDate = await item.$eval('.jeg_post_meta .jeg_meta_date a', (el: any) => (el.textContent || '').trim());
            } catch {
                postDate = '';
            }

            try {
                postDescription = await item.$eval('.jeg_post_excerpt', (el: any) => (el.textContent || '').trim());
            } catch {
                postDescription = '';
            }

            try {
                postLink = await item.$eval('.jeg_postblock_content .jeg_post_title a', (el: any) => el.getAttribute('href') || '');
            } catch {
                postLink = '';
            }

            localResults.push({
                post_date: postDate,
                post_title: postTitle,
                post_description: postDescription,
                post_link: postLink,
                post_category,
            });
        }

        return localResults;
    }

    for (const url of urls) {
        console.log('Processing', url);

        for (const lang of languages) {
            const targetFolder = lang.folder;
            const candidates = [
                url,
                `${url}?lang=${lang.code}`,
                `${url.replace(/\/$/, '')}/${lang.code}/`,
            ];

            let scraped = false;
            let langResults: any[] = [];

            for (const candidate of candidates) {
                try {
                    await page.goto(candidate, { timeout: 120_000, waitUntil: 'load' });
                } catch (e) {
                    console.warn('Failed to load candidate', candidate, e);
                    continue;
                }

                // detect html lang attribute
                let pageLang = '';
                try {
                    pageLang = await page.$eval('html', (el: any) => el.getAttribute('lang') || '');
                } catch {
                    pageLang = '';
                }

                if (pageLang && pageLang.startsWith(lang.code)) {
                    // scrape
                    langResults = await scrapeFromPage(page);
                    scraped = true;
                    console.log(`Scraped ${lang.code} from ${candidate} (${lang.folder}) - items:`, langResults.length);
                    break;
                }

                // try to find language links and click (hreflang/lang attr)
                try {
                    const selector = `a[hreflang="${lang.code}"], a[lang="${lang.code}"]`;
                    const langHandle = await page.$(selector);
                    if (langHandle) {
                        await Promise.all([
                            page.waitForNavigation({ waitUntil: 'load', timeout: 10_000 }).catch(() => null),
                            langHandle.click().catch(() => null),
                        ]);

                        try {
                            pageLang = await page.$eval('html', (el: any) => el.getAttribute('lang') || '');
                        } catch {
                            pageLang = '';
                        }

                        if (pageLang && pageLang.startsWith(lang.code)) {
                            langResults = await scrapeFromPage(page);
                            scraped = true;
                            console.log(`Scraped ${lang.code} after clicking lang link on ${candidate} (${lang.folder}) - items:`, langResults.length);
                            break;
                        }
                    }
                } catch {
                    // ignore click failures
                }
            }

            // write per-language outputs
            try {
                const langOutJsonDir = path.resolve(__dirname, '..', 'outputs', 'data', targetFolder, 'json');
                const langOutXlsxDir = path.resolve(__dirname, '..', 'outputs', 'data', targetFolder, 'xlsx');
                await mkdir(langOutJsonDir, { recursive: true });
                await mkdir(langOutXlsxDir, { recursive: true });

                const slug = url.replace(/https?:\/\//, '').replace(/[^a-z0-9]+/gi, '_').replace(/^_+|_+$/g, '').toLowerCase();

                // per-language file
                const perLangJson = path.resolve(langOutJsonDir, `output_${slug}_${targetFolder}.json`);
                await writeFile(perLangJson, JSON.stringify(langResults, null, 2), 'utf8');
                console.log('Saved per-language JSON to', perLangJson);

                // combined per-language outputs (overwrite)
                const combinedJsonPath = path.resolve(path.resolve(__dirname, '..', 'outputs', 'data', targetFolder, 'json'), 'outputs.json');
                await writeFile(combinedJsonPath, JSON.stringify(langResults, null, 2), 'utf8');

                const rows = langResults.map(r => ({
                    'Post date': r.post_date,
                    'Post title': r.post_title,
                    'Post description': r.post_description,
                    'Post link': r.post_link,
                    'Post category': r.post_category || '',
                }));

                const aoa: any[][] = [headers, ...rows.map(row => headers.map(h => row[h] ?? ''))];
                const ws = XLSX.utils.aoa_to_sheet(aoa);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
                const combinedXlsxPath = path.resolve(path.resolve(__dirname, '..', 'outputs', 'data', targetFolder, 'xlsx'), 'outputs.xlsx');
                XLSX.writeFile(wb, combinedXlsxPath);
                console.log('Saved per-language XLSX to', combinedXlsxPath);
            } catch (err) {
                console.error('Failed to write per-language outputs for', lang.folder, err);
            }
        }
    }
});