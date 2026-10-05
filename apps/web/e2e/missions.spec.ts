import { test, expect, type Page } from '@playwright/test';

async function ask(page: Page, prompt: string) {
  const local = page.getByRole('button', { name: /Try the local agent/i });
  if (await local.isVisible()) await local.click();
  await page.getByPlaceholder(/Ask ExcelAgento/i).fill(prompt);
  await page.getByRole('button', { name: /^Send$/i }).click();
  await expect(page.getByRole('button', { name: /Apply Changes/i }).last()).toBeVisible();
}
async function missions(page: Page) {
  await page.getByRole('button', { name: /Missions/i }).click();
  await expect(page.getByRole('status', { name: 'Mission storage status' })).toHaveText(
    'Mission history saved locally',
  );
}

test('prepared mission survives reload, matches restored checkpoint, reconfirms and reverses the actual worker commit', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/');
  await ask(page, 'remove duplicate rows');
  await missions(page);
  await expect(page.getByText('Ready for review')).toBeVisible();
  await page.reload();
  await expect(page.getByRole('button', { name: 'Resume review' })).toBeVisible();
  await page.getByRole('button', { name: /Back to workspace/i }).click();
  const restore = page.getByRole('button', { name: 'Restore checkpoint' });
  if (await restore.isVisible()) await restore.click();
  await missions(page);
  await page.getByRole('button', { name: 'Resume review' }).click();
  await expect(page.getByText(/engine generated a fresh preview/i)).toBeVisible();
  await page.getByRole('button', { name: /Apply Changes/i }).click();
  await expect(page.getByRole('group', { name: 'Confirm destructive change' })).toBeVisible();
  await expect(page.locator('.file-meta-dims')).toContainText('11 rows');
  await page.getByRole('button', { name: 'Yes, apply this change' }).click();
  await expect(page.locator('.file-meta-dims')).toContainText('10 rows');
  await expect(page.getByText('Task completed', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: /Undo this step/i }).click();
  await expect(page.locator('.file-meta-dims')).toContainText('11 rows');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(page.locator('.file-meta-dims')).toContainText('10 rows');
  await missions(page);
  await expect(page.getByRole('article').getByText('Completed', { exact: true })).toBeVisible();
  await expect(page.getByRole('button', { name: 'Resume review' })).toHaveCount(0);
  await page.getByRole('button', { name: 'Remove', exact: true }).click();
  await page.getByRole('button', { name: 'Delete mission data' }).click();
  await expect(page.getByRole('article')).toHaveCount(0);
  await page.getByRole('button', { name: /Back to workspace/i }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByText('Task reverted', { exact: true })).toBeVisible();
  await missions(page);
  await page.reload();
  await expect(page.getByRole('article')).toHaveCount(0);
  expect(errors).toEqual([]);
});

test('manual edits stale prepared work and offer a new request instead of unsafe resume', async ({
  page,
}) => {
  await page.goto('/');
  await ask(page, 'trim whitespace in column B');
  await missions(page);
  await page.getByRole('button', { name: /Back to workspace/i }).click();
  const cell = page.getByRole('gridcell', { name: /^B2: / });
  await cell.dblclick();
  await page.locator('#grid-cell-editor').fill('Changed customer');
  await page.locator('#grid-cell-editor').press('Enter');
  await missions(page);
  await expect(page.getByText('Needs re-check')).toBeVisible();
  await expect(page.getByRole('button', { name: 'Resume review' })).toHaveCount(0);
  await page.getByRole('button', { name: 'Replan request' }).click();
  await expect(page.getByPlaceholder(/Ask ExcelAgento/i)).toHaveValue(
    'trim whitespace in column B',
  );
});

test('responsive analyst briefing gives source ranges, chart alternatives and exact baseline comparison without a provider', async ({
  page,
}) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('/');
  await page.getByRole('button', { name: 'Insights', exact: true }).click();
  await expect(page.getByRole('heading', { name: 'Analyst briefing', exact: true })).toBeVisible();
  await expect(page.getByText('Local computation · No model required')).toBeVisible();
  await expect(page.getByRole('region', { name: 'Data quality summary' })).toBeVisible();
  await expect(page.getByRole('region', { name: 'Baseline comparison' })).toBeVisible();
  const width = await page.evaluate(() => ({
    viewport: window.innerWidth,
    scroll: document.documentElement.scrollWidth,
  }));
  expect(width.scroll).toBeLessThanOrEqual(width.viewport + 1);
});
