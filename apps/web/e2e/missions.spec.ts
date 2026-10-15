import { test, expect, type Page } from '@playwright/test';

async function ask(page: Page, prompt: string) {
  const agentTab = page.getByRole('button', { name: 'Agent', exact: true });
  if (await agentTab.isVisible()) await agentTab.click();
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

test('first-run starter carries a new user through review, apply and export', async ({ page }) => {
  await page.goto('/');
  await expect(page.getByRole('heading', { name: 'Get your first workbook result' })).toBeVisible();
  await page.getByRole('button', { name: 'Try sample cleanup' }).click();
  const composer = page.getByPlaceholder(/Ask ExcelAgento/i);
  await expect(composer).toHaveValue('remove duplicate rows');
  await page.getByRole('button', { name: /^Send$/i }).click();
  await page
    .getByRole('button', { name: /Apply Changes/i })
    .last()
    .click();
  await expect(page.getByRole('group', { name: 'Confirm destructive change' })).toBeVisible();
  await page.getByRole('button', { name: 'Yes, apply this change' }).click();
  await expect(page.getByText('Task completed', { exact: true })).toBeVisible();

  const downloadPromise = page.waitForEvent('download');
  await page
    .getByRole('button', { name: /^Export(?: \.xlsx)?$/i })
    .first()
    .click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toMatch(/\.xlsx$/i);
});

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

test('workspace chrome and verification card stay contained across desktop and phone widths', async ({
  page,
}) => {
  for (const viewport of [
    { width: 1440, height: 900 },
    { width: 1024, height: 900 },
    { width: 390, height: 844 },
    { width: 320, height: 844 },
  ]) {
    await page.setViewportSize(viewport);
    await page.goto('/');
    await expect(page.getByRole('grid')).toBeVisible();
    await expect
      .poll(() => page.evaluate(() => document.documentElement.scrollWidth))
      .toBeLessThanOrEqual(viewport.width + 1);

    const resourceRight = await page.locator('.studio-resource-nav').evaluate((element) => {
      const rect = element.getBoundingClientRect();
      return rect.right;
    });
    expect(resourceRight).toBeLessThanOrEqual(viewport.width + 1);
  }

  await page.setViewportSize({ width: 1440, height: 900 });
  await page.goto('/#/verify-email');
  const verificationCard = page.locator('.verify-email-card');
  await expect(verificationCard).toBeVisible();
  await expect.poll(() => verificationCard.boundingBox()).toMatchObject({ width: 520 });
});

test('keeps the workbook usable on a compact laptop and removes sheets safely', async ({
  page,
}) => {
  await page.setViewportSize({ width: 1366, height: 768 });
  await page.goto('/');

  await expect(page.locator('.studio-mobile-tabs')).toBeVisible();
  await expect(page.getByRole('grid')).toBeVisible();
  await expect(page.locator('.studio-data-area')).toHaveCSS('min-width', '0px');

  page.on('dialog', (dialog) => void dialog.accept());
  await page.getByRole('button', { name: 'Remove sheet Summary & Metrics' }).click();

  await expect(page.getByRole('button', { name: 'Remove sheet Summary & Metrics' })).toHaveCount(0);
  await expect(page.getByRole('grid')).toHaveAttribute(
    'aria-label',
    'Orders & Deliveries spreadsheet',
  );

  await page.getByRole('button', { name: 'Remove sheet Orders & Deliveries' }).click();
  await expect(page.locator('.toast-host')).toContainText(
    'A workbook must keep at least one sheet.',
  );
});
