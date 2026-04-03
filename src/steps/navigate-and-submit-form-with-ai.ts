import * as crypto from 'crypto';
import { BaseStep, ExpectedRecord, Field, StepInterface } from '../core/base-step';
import { Step, RunStepResponse, FieldDefinition, StepDefinition, RecordDefinition, StepRecord } from '../proto/cog_pb';
import { AiFormCache } from '../client/ai-form-cache';
import { AiFormFill, AiFillAction } from '../client/mixins/ai-form-fill';

export class NavigateAndSubmitFormWithAI extends BaseStep implements StepInterface {

  protected stepName: string = 'Navigate and submit form with AI';
  protected stepExpression: string = 'navigate and fill a form at (?<webPageUrl>.+) using AI';
  protected stepType: StepDefinition.Type = StepDefinition.Type.ACTION;
  protected actionList: string[] = ['navigate'];
  protected targetObject: string = 'Navigate and fill out form with AI';

  protected expectedFields: Field[] = [
    {
      field: 'webPageUrl',
      type: FieldDefinition.Type.URL,
      description: 'Page URL',
    },
    {
      field: 'fieldOverrides',
      type: FieldDefinition.Type.MAP,
      description: 'Optional map of specific field values to enforce (matched by label, name, or placeholder), e.g. Country → U.S.A., State → Georgia',
      optionality: FieldDefinition.Optionality.OPTIONAL,
    },
    {
      field: 'maxAttempts',
      type: FieldDefinition.Type.NUMERIC,
      description: 'Max AI retry attempts on failure (default 2, max 3)',
      optionality: FieldDefinition.Optionality.OPTIONAL,
    },
    {
      field: 'cacheStrategy',
      type: FieldDefinition.Type.STRING,
      description: '"hash" (default — Redis cache, re-runs AI only when form changes), "promote" (rewrites scenario with concrete steps on first success), "always" (always calls AI)',
      optionality: FieldDefinition.Optionality.OPTIONAL,
    },
  ];

  protected expectedRecords: ExpectedRecord[] = [
    {
      id: 'formFill',
      type: RecordDefinition.Type.KEYVALUE,
      fields: [
        { field: 'url', type: FieldDefinition.Type.STRING, description: 'URL navigated to' },
        { field: 'actionsCount', type: FieldDefinition.Type.NUMERIC, description: 'Number of fill actions executed' },
        { field: 'attempts', type: FieldDefinition.Type.NUMERIC, description: 'Number of AI attempts made (0 if served from cache)' },
        { field: 'fromCache', type: FieldDefinition.Type.BOOLEAN, description: 'Whether fill actions were served from the Redis cache' },
        { field: 'cacheStrategy', type: FieldDefinition.Type.STRING, description: 'Cache strategy used' },
      ],
      dynamicFields: true,
    },
  ];

  async executeStep(step: Step): Promise<RunStepResponse> {
    const stepData: any = step.getData().toJavaScript();
    const url: string = stepData.webPageUrl;
    const cacheStrategy: string = stepData.cacheStrategy || 'hash';
    const maxAttempts: number = Math.min(Math.max(parseInt(stepData.maxAttempts) || 2, 1), 3);
    const stepOrder: number = stepData['__stepOrder'] || 1;

    // MAP fields are delivered as plain JS objects by the proto Struct deserializer
    const fieldOverrides: Record<string, string> = (stepData.fieldOverrides && typeof stepData.fieldOverrides === 'object')
      ? stepData.fieldOverrides
      : {};

    const requestorId: string = (this.client.idMap && this.client.idMap.requestorId) || 'unknown';

    console.log(`[AI-Step] NavigateAndSubmitFormWithAI — url=${url} strategy=${cacheStrategy} maxAttempts=${maxAttempts} requestorId=${requestorId}`);

    // ── Step 1: Navigate ────────────────────────────────────────────────────────
    // Start the 'time' label expected by basic-interaction.ts checkpoints.
    // Do NOT call timeEnd here — the submit click triggers a second navigation
    // that also uses timeLog('time'), so the timer must stay alive for the
    // lifetime of the step.
    console.time('time');
    try {
      await this.client.navigateToUrl(
        url,
        stepData.throttle || false,
        stepData.maxInflightRequests || 0,
      );
    } catch (navErr) {
      console.log(`[AI-Step] Navigation failed — ${url}: ${navErr.toString()}`);
      const screenshot = await this.safeCapture();
      return this.error(
        'Failed to navigate to %s: %s',
        [url, navErr.toString()],
        screenshot ? [this.binary('screenshot', 'Screenshot', 'image/jpeg', screenshot)] : [],
      );
    }

    // ── Step 2: Extract rendered form HTML (for AI) and a stable structural
    //            fingerprint (for cache keying). The fingerprint omits volatile
    //            values like CSRF tokens and session IDs that change each page
    //            load but don't reflect a real form change. ──────────────────────
    let formHtml = '';
    try {
      formHtml = await this.client.client.evaluate(() => {
        const forms = Array.from(document.querySelectorAll('form'));
        return forms.length > 0
          ? forms.map((f: Element) => (f as HTMLElement).outerHTML).join('\n')
          : document.body.innerHTML;
      }) as string;
    } catch (htmlErr) {
      console.log(`[AI-Step] HTML extraction failed (non-fatal, AI will rely on screenshot) — ${htmlErr.toString()}`);
    }

    // Build a structural fingerprint from field names/types/ids/options only.
    // This is stable across CSRF token rotations, session IDs, and other
    // per-load dynamic values that don't change the form's actual structure.
    let formStructure = '';
    try {
      formStructure = await this.client.client.evaluate(() => {
        const fields: object[] = [];
        document.querySelectorAll('form input, form select, form textarea, form button[type="submit"]').forEach((el) => {
          const entry: any = {
            tag: el.tagName.toLowerCase(),
            name: (el as HTMLInputElement).name || '',
            id: el.id || '',
            type: (el as HTMLInputElement).type || '',
          };
          // Include the set of option values for selects (order-insensitive sort for stability)
          if (el.tagName === 'SELECT') {
            entry.options = Array.from((el as HTMLSelectElement).options)
              .map((o) => o.value)
              .filter(Boolean)
              .sort();
          }
          // Skip hidden inputs — common vector for CSRF tokens
          if (entry.tag === 'input' && entry.type === 'hidden') return;
          fields.push(entry);
        });
        return JSON.stringify(fields);
      }) as string;
    } catch (_) {
      // Fallback: hash the raw HTML if evaluate fails
      formStructure = formHtml;
    }

    const formHash = crypto.createHash('sha256').update(formStructure).digest('hex');

    // ── Step 3: Redis cache check (hash strategy only) ──────────────────────────
    if (cacheStrategy === 'hash') {
      const cache = new AiFormCache(this.client.redisClient);
      const cached = await cache.get(requestorId, url);

      if (cached) {
        if (cached.formHash === formHash) {
          const nonClickActions = cached.fillActions.filter(a => a.inputType !== 'click');
          const clickActions = cached.fillActions.filter(a => a.inputType === 'click');
          const overrideCount = Object.keys(fieldOverrides).length;
          console.log(`[AI-Step] Cache HIT — executing ${cached.fillActions.length} cached actions for ${url}${overrideCount > 0 ? ` (applying ${overrideCount} fieldOverride(s) via DOM)` : ''}`);
          try {
            await this.executeFillActions(nonClickActions);
            if (overrideCount > 0) {
              await this.applyOverridesToPage(fieldOverrides);
            }
            await this.executeFillActions(clickActions);
            await this.sleep(1500);
            const screenshot = await this.safeCapture();
            const records = this.buildRecords(url, cached.fillActions.length, 0, true, cacheStrategy, cached.fillActions, screenshot, stepOrder);
            return this.pass(
              'Successfully filled out and submitted the form at %s (served from Redis cache)',
              [url],
              records,
            );
          } catch (execErr) {
            console.log(`[AI-Step] Cached fill failed — falling through to AI: ${execErr.toString()}`);
          }
        } else {
          console.log(`[AI-Step] Cache STALE — form structure changed for ${url}, re-running AI`);
        }
      } else {
        console.log(`[AI-Step] Cache MISS — no entry for ${url}`);
      }
    }

    // ── Step 4: AI fill loop ─────────────────────────────────────────────────────
    let aiHelper: AiFormFill | null = null;
    try {
      aiHelper = new AiFormFill();
    } catch (configErr) {
      console.log(`[AI-Step] Azure OpenAI not configured: ${configErr.toString()}`);
      const screenshot = await this.safeCapture();
      return this.error(
        'Azure OpenAI is not configured: %s',
        [configErr.toString()],
        screenshot ? [this.binary('screenshot', 'Screenshot', 'image/jpeg', screenshot)] : [],
      );
    }

    const initialScreenshot = await this.safeCaptureBase64();
    const preSubmitUrl = await this.safeGetUrl();

    let messages: any[] = [];
    let fillActions: AiFillAction[] | null = null;
    let lastError = '';
    let lastScreenshot: any = null;

    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      try {
        // On retry: check whether the previous attempt's submit actually navigated
        // the page away before re-filling from scratch. If it did, we're done.
        if (attempt > 1) {
          const alreadySubmitted = await this.detectSuccess(preSubmitUrl);
          if (alreadySubmitted) {
            console.log(`[AI-Step] Late success detected after attempt ${attempt - 1} for ${url}`);
            lastScreenshot = await this.safeCapture();
            const records = this.buildRecords(url, fillActions!.length, attempt - 1, false, cacheStrategy, fillActions!, lastScreenshot, stepOrder);
            if (cacheStrategy === 'hash') {
              const cache = new AiFormCache(this.client.redisClient);
              await cache.set(requestorId, url, formHash, fillActions!);
            }
            return this.pass(
              'Successfully filled out and submitted the form at %s after %d AI attempt(s)',
              [url, attempt - 1],
              records,
            );
          }
          const retryScreenshot = await this.safeCaptureBase64();
          messages.push(aiHelper.buildRetryMessage(retryScreenshot, lastError));
        }

        console.log(`[AI-Step] Calling AI — attempt ${attempt}/${maxAttempts} for ${url}`);
        const result = await aiHelper.getFillActions(initialScreenshot, formHtml, fieldOverrides, messages);
        messages = result.messages;
        fillActions = result.actions;

        await this.executeFillActions(fillActions);
        await this.sleep(1500);

        const submitted = await this.detectSuccess(preSubmitUrl);

        if (submitted) {
          if (cacheStrategy === 'hash') {
            const cache = new AiFormCache(this.client.redisClient);
            await cache.set(requestorId, url, formHash, fillActions);
          }
          lastScreenshot = await this.safeCapture();
          const records = this.buildRecords(url, fillActions.length, attempt, false, cacheStrategy, fillActions, lastScreenshot, stepOrder);
          console.log(`[AI-Step] Form submitted successfully on attempt ${attempt} for ${url}`);
          return this.pass(
            'Successfully filled out and submitted the form at %s after %d AI attempt(s)',
            [url, attempt],
            records,
          );
        }

        lastError = await this.extractErrorText();
        console.log(`[AI-Step] Attempt ${attempt} did not detect submission — page errors: "${lastError || '(none)'}"`);
        lastScreenshot = await this.safeCapture();

      } catch (attemptErr) {
        lastError = attemptErr.toString();
        lastScreenshot = await this.safeCapture();
        console.log(`[AI-Step] Attempt ${attempt} error: ${lastError}`);
      }
    }
    console.log(`[AI-Step] All ${maxAttempts} attempt(s) exhausted for ${url}`);

    // ── All attempts exhausted ──────────────────────────────────────────────────
    const errorRecords: StepRecord[] = [];
    if (lastScreenshot) {
      errorRecords.push(this.binary('screenshot', 'Final Screenshot', 'image/jpeg', lastScreenshot));
    }
    errorRecords.push(this.keyValue('formFill', 'Form Fill Failed', {
      url,
      actionsCount: fillActions ? fillActions.length : 0,
      attempts: maxAttempts,
      fromCache: false,
      cacheStrategy,
      lastError,
    }));

    return this.error(
      'Failed to fill and submit the form at %s after %d attempt(s). Last error: %s',
      [url, maxAttempts, lastError],
      errorRecords,
    );
  }

  // ── Private helpers ──────────────────────────────────────────────────────────

  /**
   * For each override key, finds the matching form field on the live page via
   * DOM inspection (checking name, id, placeholder, and label text) and
   * re-fills it with the override value. This is called after the cached fill
   * actions have already run, so it acts as a guaranteed second-pass correction
   * that is independent of the selector format the AI happened to use.
   */
  private async applyOverridesToPage(overrides: Record<string, string>): Promise<void> {
    for (const [key, value] of Object.entries(overrides)) {
      try {
        const selector = await this.resolveFieldSelectorByKey(key);
        if (selector) {
          // Clear any value the cached fill already placed in this field.
          // fillOutField uses page.type() which appends, so we must wipe first.
          await this.client.client.evaluate((sel: string) => {
            const el = document.querySelector(sel) as HTMLInputElement;
            if (el && 'value' in el) {
              el.value = '';
              el.dispatchEvent(new Event('input', { bubbles: true }));
              el.dispatchEvent(new Event('change', { bubbles: true }));
            }
          }, selector);
          await this.client.fillOutField(selector, value);
        } else {
          console.log(`[AI-Step] Override: no field matched key "${key}" on the page — skipping`);
        }
      } catch (err) {
        console.log(`[AI-Step] Override fill failed for key "${key}": ${err}`);
      }
    }
  }

  /**
   * Evaluates the page DOM to find a form field matching the given key.
   * Checks (in order): name attribute, id attribute, placeholder text, and
   * associated label text — all compared case-insensitively with punctuation
   * stripped. Returns a minimal CSS selector for the matched element, or null.
   */
  private async resolveFieldSelectorByKey(key: string): Promise<string | null> {
    const normalizedKey = key.toLowerCase().replace(/[^a-z0-9]/g, '');
    return await this.client.client.evaluate((normKey: string) => {
      const norm = (s: string) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      const matches = (a: string) => !!a && (a === normKey || a.includes(normKey) || normKey.includes(a));
      const fields = Array.from(
        document.querySelectorAll('form input:not([type="hidden"]), form select, form textarea'),
      ) as HTMLInputElement[];

      for (const el of fields) {
        let labelText = '';
        if (el.id) {
          const label = document.querySelector(`label[for="${el.id}"]`);
          if (label) labelText = norm(label.textContent || '');
        }

        if (
          matches(norm(el.name)) ||
          matches(norm(el.id)) ||
          matches(norm((el as any).placeholder || '')) ||
          matches(labelText)
        ) {
          if (el.name) return `[name="${el.name}"]`;
          if (el.id) return `#${el.id}`;
        }
      }
      return null;
    }, normalizedKey) as string | null;
  }

  private async executeFillActions(actions: AiFillAction[]): Promise<void> {
    for (let i = 0; i < actions.length; i++) {
      const action = actions[i];
      try {
        if (action.inputType === 'click') {
          await this.client.submitFormByClickingButton(action.selector);
        } else {
          await this.client.fillOutField(action.selector, action.value);
        }
      } catch (actionErr) {
        if (action.inputType === 'click') {
          // submitFormByClickingButton can throw even when the page navigated successfully
          // (e.g. "Submit button still there" race during navigation). Don't re-throw —
          // let detectSuccess() determine the real outcome from page state.
          console.log(`[AI-Step] Click action threw (may be a navigation race) — selector="${action.selector}": ${actionErr.toString()}`);
        } else {
          console.log(`[AI-Step] Fill action failed — selector="${action.selector}" inputType=${action.inputType}: ${actionErr.toString()}`);
          throw actionErr;
        }
      }
    }
  }

  private async detectSuccess(preSubmitUrl: string): Promise<boolean> {
    try {
      const currentUrl = await this.client.client.url();
      if (currentUrl !== preSubmitUrl) return true;

      const formGone: boolean = await this.client.client.evaluate(
        () => document.querySelectorAll('form').length === 0,
      ) as boolean;
      if (formGone) return true;

      const hasSuccessMessage: boolean = await this.client.client.evaluate(() => {
        const text = (document.body as HTMLElement).innerText.toLowerCase();
        return ['thank you', 'thanks!', 'success', 'submitted', 'received', 'confirmed']
          .some((p) => text.includes(p));
      }) as boolean;

      return hasSuccessMessage;
    } catch (_) {
      return false;
    }
  }

  private async extractErrorText(): Promise<string> {
    try {
      const text: string = await this.client.client.evaluate(() => {
        const selectors = [
          '.error', '.errors', '.alert-danger', '.field-error',
          '[class*="error"]', '[class*="invalid"]', '[aria-invalid="true"]',
        ];
        const found: string[] = [];
        for (const sel of selectors) {
          document.querySelectorAll(sel).forEach((el) => {
            const t = ((el as HTMLElement).innerText || '').trim();
            if (t) found.push(t);
          });
        }
        return found.slice(0, 5).join('; ');
      }) as string;
      return text;
    } catch (_) {
      return '';
    }
  }

  private async safeCapture(): Promise<any> {
    try {
      return await this.client.safeScreenshot({ type: 'jpeg', encoding: 'binary', quality: 60 });
    } catch (_) {
      return null;
    }
  }

  private async safeCaptureBase64(): Promise<string> {
    try {
      return (await this.client.safeScreenshot({ type: 'jpeg', encoding: 'base64', quality: 60 })) as string;
    } catch (_) {
      return '';
    }
  }

  private async safeGetUrl(): Promise<string> {
    try {
      return await this.client.client.url();
    } catch (_) {
      return '';
    }
  }

  private sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  private buildRecords(
    url: string,
    actionsCount: number,
    attempts: number,
    fromCache: boolean,
    cacheStrategy: string,
    fillActions: AiFillAction[],
    screenshot: any,
    stepOrder: number,
  ): StepRecord[] {
    const records: StepRecord[] = [];

    if (screenshot) {
      records.push(this.binary('screenshot', 'Screenshot', 'image/jpeg', screenshot));
    }

    const summaryData = { url, actionsCount, attempts, fromCache, cacheStrategy };
    records.push(this.keyValue('formFill', 'Form Fill Summary', summaryData));
    records.push(this.keyValue(`formFill.${stepOrder}`, `Form Fill Summary from Step ${stepOrder}`, summaryData));

    // For promote strategy: include the concrete replacement steps so the app can
    // rewrite this scenario with deterministic steps on subsequent runs.
    if (cacheStrategy === 'promote' && fillActions.length > 0) {
      const promotedSteps = this.buildPromotedSteps(url, fillActions);
      // The 'exposeOnPass:' prefix tells the cog-mechanism's StepResponse.fromProto
      // to keep this record in the run log even when the step outcome is Passed.
      // Without it, all non-allowlisted records are stripped before the log is saved.
      records.push(this.keyValue('exposeOnPass:promotedSteps', 'Promoted Concrete Steps', {
        stepsJson: JSON.stringify(promotedSteps),
      }));
    }

    return records;
  }

  /**
   * Builds the concrete step definitions that replace this AI step when
   * cacheStrategy === "promote". The app's run completion handler detects the
   * "promotedSteps" record and rewrites the scenario definition.
   */
  private buildPromotedSteps(url: string, fillActions: AiFillAction[]): any[] {
    const steps: any[] = [
      {
        stepId: 'NavigateToPage',
        name: 'Navigate to a webpage',
        cog: 'automatoninc/web',
        data: { webPageUrl: url },
      },
    ];

    for (const action of fillActions) {
      if (action.inputType === 'click') {
        steps.push({
          stepId: 'SubmitFormByClickingButton',
          name: 'Submit a form by clicking a button',
          cog: 'automatoninc/web',
          data: { domQuerySelector: action.selector },
        });
      } else {
        steps.push({
          stepId: 'EnterValueIntoField',
          name: 'Fill out a form field',
          cog: 'automatoninc/web',
          data: {
            domQuerySelector: action.selector,
            value: action.value,
          },
        });
      }
    }

    return steps;
  }
}

export { NavigateAndSubmitFormWithAI as Step };
