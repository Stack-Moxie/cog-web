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
      description: '"hash" (default — Redis cache, re-runs AI only when form changes) or "promote" (rewrites scenario with concrete steps on first success)',
      optionality: FieldDefinition.Optionality.OPTIONAL,
    },
    {
      field: 'userHint',
      type: FieldDefinition.Type.STRING,
      description: 'Optional additional instructions for the AI about this specific form (e.g. "Click the Contact Us tab first to reveal the form" or "Select Country before the State dropdown appears"). Must relate to form interaction only.',
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
    const maxAttempts: number = Math.min(Math.max(parseInt(stepData.maxAttempts) || 2, 1), 3);
    const stepOrder: number = stepData['__stepOrder'] || 1;

    // Only 'hash' and 'promote' are supported. 'always' was removed to prevent
    // unbounded AI spend; anything else is silently normalised to 'hash'.
    const rawStrategy: string = stepData.cacheStrategy || 'hash';
    const cacheStrategy: string = ['hash', 'promote'].includes(rawStrategy) ? rawStrategy : 'hash';
    if (rawStrategy !== cacheStrategy) {
      console.log(`[AI-Step] Unsupported cacheStrategy "${rawStrategy}" — falling back to "hash"`);
    }

    // MAP fields are delivered as plain JS objects by the proto Struct deserializer
    const fieldOverrides: Record<string, string> = (stepData.fieldOverrides && typeof stepData.fieldOverrides === 'object')
      ? stepData.fieldOverrides
      : {};

    // Sanitize the user-supplied hint before it reaches the AI.
    const userHint: string = AiFormFill.sanitizeUserHint(stepData.userHint || '');
    if (stepData.userHint && !userHint) {
      console.log('[AI-Step] userHint was provided but contained disallowed content and was ignored.');
    }

    const requestorId: string = (this.client.idMap && this.client.idMap.requestorId) || 'unknown';

    console.log(`[AI-Step] NavigateAndSubmitFormWithAI — url=${url} strategy=${cacheStrategy} maxAttempts=${maxAttempts} requestorId=${requestorId}${userHint ? ' (userHint provided)' : ''}`);

    // ── Step 1: Navigate ────────────────────────────────────────────────────────
    // Start the 'time' label expected by basic-interaction.ts checkpoints.
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

    // ── Step 2a: Dismiss cookie/consent banners ──────────────────────────────────
    // Cookie consent overlays block scrolling and interaction on many pages.
    // Attempt to click common "accept" buttons before doing anything else.
    await this.dismissCookieBanners();

    // ── Step 2b: Scroll to 75% to trigger intersection-observer lazy loading ─────
    // We intentionally stay at 75% here (not scrolling back yet) so that any
    // revelation AI call can capture the HTML/screenshot while the lazy-loaded
    // content (e.g. the location dropdown) is still in the DOM and viewport.
    // The scroll-back to 0% happens after the HTML snapshot is taken below.
    try {
      await this.client.scrollTo(75, '%');
      await this.sleep(2000); // allow dynamic content to fully render after scroll
    } catch (_) {}
    // Pre-capture the HTML and screenshot at 75% depth, BEFORE scrolling back.
    // Pages that use virtual/lazy rendering remove off-screen elements from the
    // DOM when scrolled away, so this is the only reliable window to capture
    // the content that the revelation AI needs to identify the correct selectors.
    let preRevealHtml = '';
    let preRevealScreenshot = '';
    if (userHint) {
      try {
        preRevealHtml = await this.captureFullPageHtml();
      } catch (_) {}
      preRevealScreenshot = await this.safeCaptureBase64();
      // Log a snippet so engineers can verify the AI has the right DOM context.
      console.log(`[AI-Step] preRevealHtml length=${preRevealHtml.length} iframes=${(preRevealHtml.match(/<iframe/gi) || []).length} (first 800 chars): ${preRevealHtml.substring(0, 800)}`);
    }
    // Now scroll back to top so revelation/fill actions start from a known position.
    try {
      await this.client.scrollTo(0, '%');
      await this.sleep(500);
    } catch (_) {}

    // ── Step 3: Configure AI early (needed for both reveal and fill phases) ─────
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

    // ── Step 4: Fetch cache entry early (for reveal actions) ─────────────────────
    const cache = new AiFormCache(this.client.redisClient);
    const cacheEntry = cacheStrategy === 'hash' ? await cache.get(requestorId, url) : null;

    // ── Step 5: Revelation phase (hint-triggered) ────────────────────────────────
    // If a userHint is provided, execute actions that reveal/load the form before
    // we take the main snapshot. Cached reveal actions skip the AI call entirely.
    let revealActions: AiFillAction[] = [];
    if (userHint) {
      if (cacheEntry?.revealActions && cacheEntry.revealActions.length > 0) {
        revealActions = cacheEntry.revealActions as AiFillAction[];
        console.log(`[AI-Step] Using ${revealActions.length} cached reveal action(s) for ${url}`);
        await this.executeFillActions(revealActions, true); // return value intentionally ignored for cached path
        const maxWait = Math.max(2000, ...revealActions.map(a => a.waitAfter || 0));
        await this.sleep(maxWait);
      } else {
        // Multi-round revelation: custom dropdowns need "open trigger → click option"
        // which requires two AI passes because the options only appear in the DOM
        // after the trigger is clicked (first round).
        const MAX_REVEAL_ROUNDS = 2;
        for (let round = 0; round < MAX_REVEAL_ROUNDS; round++) {
          // Round 1 uses the pre-captured 75%-scroll snapshot (captured before the
          // scroll-back above). This ensures the AI sees the dropdown while it is
          // still in the DOM and visible, rather than a post-scroll 0% snapshot
          // where lazy-rendered elements may have been unmounted.
          // Subsequent rounds re-capture at the current (post-action) page state.
          let revealPageHtml: string;
          let revealScreenshot: string;
          if (round === 0) {
            revealPageHtml = preRevealHtml;
            revealScreenshot = preRevealScreenshot;
          } else {
            revealPageHtml = '';
            try {
              revealPageHtml = await this.captureFullPageHtml();
            } catch (_) {}
            revealScreenshot = await this.safeCaptureBase64();
          }

          console.log(`[AI-Step] Calling reveal AI (round ${round + 1}/${MAX_REVEAL_ROUNDS}) for ${url}`);

          let roundActions: AiFillAction[] = [];
          try {
            roundActions = await aiHelper.getRevealActions(revealScreenshot, revealPageHtml, userHint);
          } catch (revealErr) {
            console.log(`[AI-Step] Reveal AI round ${round + 1} failed (non-fatal): ${revealErr.toString()}`);
            break;
          }

          if (roundActions.length === 0) {
            console.log(`[AI-Step] Reveal AI returned no actions (round ${round + 1}) — form appears visible`);
            break;
          }

          // Accumulate actions from all rounds; the full sequence is cached so that
          // replays can reproduce the same multi-step revelation without calling the AI.
          revealActions.push(...roundActions);

          console.log(`[AI-Step] Executing ${roundActions.length} reveal action(s) (round ${round + 1})`);
          const anyRevealFailed = await this.executeFillActions(roundActions, true);
          const maxWait = Math.max(2000, ...roundActions.map(a => a.waitAfter || 0));
          await this.sleep(maxWait);

          // Only proceed to the next round if actions failed — successful actions (e.g. a
          // native <select>) mean the form is likely already loading, and an extra round
          // would just duplicate work and waste AI tokens.
          if (!anyRevealFailed) break;
        }
      }
    }

    // ── Step 6: Extract rendered form HTML + structural fingerprint ──────────────
    // Both are taken AFTER revelation so they reflect the fully-loaded form state.
    //
    // Strategy: surgically collect only form-relevant HTML to avoid accidentally
    // capturing unrelated page content (e.g. "Complaint reporting" tiles) that can
    // trigger Azure OpenAI's content management policy.
    //   1. For every <select> that lives OUTSIDE a <form>, walk up 3 levels to
    //      capture just the field's label+wrapper — not the entire section.
    //   2. Append every <form> element (inline <style> stripped).
    // This gives the AI all field IDs/labels/options without surrounding page noise.
    let formHtml = '';
    try {
      formHtml = await this.client.client.evaluate(() => {
        function stripStyles(el: Element): Element {
          const clone = el.cloneNode(true) as Element;
          clone.querySelectorAll('style').forEach(s => s.remove());
          return clone;
        }

        // Is this element inside a <form>?
        function insideForm(el: Element): boolean {
          let p: Element | null = el.parentElement;
          while (p) {
            if (p.tagName === 'FORM') return true;
            p = p.parentElement;
          }
          return false;
        }

        const parts: string[] = [];

        // Part 1: selects outside any <form> — include their immediate wrapper
        // context (up to 3 ancestor levels) so labels are visible to the AI.
        Array.from(document.querySelectorAll('select'))
          .filter(sel => !insideForm(sel))
          .forEach(sel => {
            let container: Element = sel;
            for (let i = 0; i < 3; i++) {
              if (!container.parentElement || container.parentElement === document.body) break;
              container = container.parentElement;
            }
            parts.push(stripStyles(container).outerHTML);
          });

        // Part 2: all <form> elements, styles stripped
        Array.from(document.querySelectorAll('form'))
          .forEach(f => parts.push(stripStyles(f).outerHTML));

        return parts.length > 0 ? parts.join('\n') : document.body.innerHTML;
      }) as string;
      console.log(`[AI-Step] formHtml length=${formHtml.length} (first 400 chars): ${formHtml.substring(0, 400)}`);
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
      formStructure = formHtml;
    }

    const formHash = crypto.createHash('sha256').update(formStructure).digest('hex');

    // ── Step 7: Redis cache check (hash strategy) ────────────────────────────────
    // cacheEntry was already fetched in Step 4. We only need the hash comparison now.
    if (cacheStrategy === 'hash' && cacheEntry) {
      if (cacheEntry.formHash === formHash) {
        const nonClickActions = cacheEntry.fillActions.filter(a => a.inputType !== 'click');
        const clickActions = cacheEntry.fillActions.filter(a => a.inputType === 'click');
        const overrideCount = Object.keys(fieldOverrides).length;
        console.log(`[AI-Step] Cache HIT — executing ${cacheEntry.fillActions.length} cached actions for ${url}${overrideCount > 0 ? ` (applying ${overrideCount} fieldOverride(s) via DOM)` : ''}`);
        try {
          await this.executeFillActions(nonClickActions);
          if (overrideCount > 0) {
            await this.applyOverridesToPage(fieldOverrides);
          }
          await this.executeFillActions(clickActions);
          await this.sleep(1500);
          const screenshot = await this.safeCapture();
          const records = this.buildRecords(url, cacheEntry.fillActions.length, 0, true, cacheStrategy, revealActions, cacheEntry.fillActions as AiFillAction[], screenshot, stepOrder);
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
    } else if (!cacheEntry) {
      console.log(`[AI-Step] Cache MISS — no entry for ${url}`);
    }

    // ── Step 8: AI fill loop ─────────────────────────────────────────────────────
    // Take the initial screenshot AFTER revelation so the AI sees the loaded form.
    const initialScreenshot = await this.safeCaptureBase64();
    const preSubmitUrl = await this.safeGetUrl();

    let messages: any[] = [];
    let fillActions: AiFillAction[] | null = null;
    let lastError = '';
    let lastScreenshot: any = null;

    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      try {
        if (attempt > 1) {
          const alreadySubmitted = await this.detectSuccess(preSubmitUrl);
          if (alreadySubmitted) {
            console.log(`[AI-Step] Late success detected after attempt ${attempt - 1} for ${url}`);
            lastScreenshot = await this.safeCapture();
            const records = this.buildRecords(url, fillActions!.length, attempt - 1, false, cacheStrategy, revealActions, fillActions!, lastScreenshot, stepOrder);
            if (cacheStrategy === 'hash') {
              await cache.set(requestorId, url, formHash, fillActions!, revealActions);
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
        const result = await aiHelper.getFillActions(initialScreenshot, formHtml, fieldOverrides, messages, userHint);
        messages = result.messages;
        fillActions = result.actions;

        // Separate submit-click actions from field-fill actions so we can run a
        // progressive scan (for conditional/cascading fields) before submitting.
        const submitClickActions = fillActions.filter(a => a.inputType === 'click');
        const fieldFillActions = fillActions.filter(a => a.inputType !== 'click');

        const anyFillFailed = await this.executeFillActions(fieldFillActions);

        // Progressive field scan: some forms reveal additional required fields
        // only after certain selections are made (e.g. a "Customer Inquiry Type"
        // dropdown that appears after "Are you a Roche customer?" = "Yes").
        // Also runs when any fill action failed, since the field may just have
        // not been in the DOM yet (conditional loading).
        await this.fillProgressiveFields(aiHelper, messages, fieldOverrides, url);

        if (anyFillFailed) {
          console.log(`[AI-Step] One or more fill actions failed on attempt ${attempt} — progressive scan ran to compensate`);
        }

        // Final DOM-based sweep just before submit: repairs any selects/inputs
        // that the AI corrupted with invalid values during the progressive rounds.
        await this.fillEmptyRequiredSelects(fieldOverrides);
        await this.fillEmptyRequiredInputs(fieldOverrides);

        await this.executeFillActions(submitClickActions);
        await this.sleep(1500);

        const submitted = await this.detectSuccess(preSubmitUrl);

        if (submitted) {
          if (cacheStrategy === 'hash') {
            await cache.set(requestorId, url, formHash, fillActions, revealActions);
          }
          lastScreenshot = await this.safeCapture();
          const records = this.buildRecords(url, fillActions.length, attempt, false, cacheStrategy, revealActions, fillActions, lastScreenshot, stepOrder);
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

  /**
   * Executes a list of fill/reveal actions in order.
   *
   * @param actions       The actions to execute.
   * @param isRevealPhase When true, "click" actions use clickElement (non-submit clicks
   *                      that reveal or navigate to a form), and errors are swallowed
   *                      so that a failed reveal step doesn't abort the whole run.
   *                      When false (default), "click" is treated as a form submit.
   * @returns             True if any action threw an error (reveal phase only), false if all succeeded.
   */
  private async executeFillActions(actions: AiFillAction[], isRevealPhase: boolean = false): Promise<boolean> {
    let anyFailed = false;
    for (let i = 0; i < actions.length; i++) {
      const action = actions[i];
      const phase = isRevealPhase ? 'reveal' : 'fill';
      console.log(`[AI-Step] [${phase}] action ${i + 1}/${actions.length}: type=${action.inputType} selector="${action.selector}" value="${action.value}"${action.waitAfter ? ` waitAfter=${action.waitAfter}` : ''}`);
      try {
        if (action.inputType === 'selectCustomDropdown') {
          await this.selectCustomDropdown(action.selector, action.value);
        } else if (action.inputType === 'focusFrame') {
          await this.client.focusFrame(action.selector);
        } else if (action.inputType === 'scroll') {
          // Scroll the target element into the viewport to trigger lazy loading.
          try {
            await this.client.client.evaluate((sel: string) => {
              const el = document.querySelector(sel);
              if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }, action.selector);
          } catch (_) {
            // Non-fatal — page may not have the element yet
          }
          await this.sleep(500);
        } else if (action.inputType === 'click') {
          if (isRevealPhase) {
            // Scroll the target into the viewport first — after smartScrollPage returns to the
            // top, elements further down the page are below the fold and Puppeteer's
            // clickElement will throw "Element may not be visible or clickable".
            try {
              await this.client.client.evaluate((sel: string) => {
                const el = document.querySelector(sel);
                if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
              }, action.selector);
              await this.sleep(600);
            } catch (_) {
              // Non-fatal — proceed even if scroll-into-view fails
            }
            await this.client.clickElement(action.selector);
          } else {
            // Form-submit click
            await this.client.submitFormByClickingButton(action.selector);
          }
        } else {
          // text / select / checkbox / radio — clear before filling to prevent
          // page.type() from appending to existing values on retries.

          // Normalise selector: the reveal AI sometimes generates a child <option>
          // selector (e.g. "select#foo option[value='X']") instead of targeting the
          // <select> element directly.  Strip the option part so that getFieldMethod
          // can find the real element and Puppeteer's frame.select() can be used.
          const normalizedSelector = isRevealPhase
            ? action.selector.replace(/\s+option(\[.*?\])?$/i, '').trim()
            : action.selector;

          // Idempotency guard for <select> elements: if the field is already set
          // to the target value, skip the clear+fill cycle entirely.  Re-setting a
          // select triggers the page's change-event cascade (e.g. re-selecting
          // "United States" resets the dependent customer-type and inquiry-type
          // dropdowns, which then disappear from the DOM).
          if (action.inputType === 'select') {
            try {
              const selectState = await this.client.client.evaluate((sel: string, targetVal: string) => {
                const el = document.querySelector(sel) as HTMLSelectElement | null;
                if (!el) return { currentVal: null, optionExists: false };
                const currentVal = el.value;
                // Check both value attribute and visible text of each option
                const tl = targetVal.toLowerCase();
                const optionExists = Array.from(el.options).some(
                  o => o.value.toLowerCase() === tl || o.text.trim().toLowerCase() === tl,
                );
                return { currentVal, optionExists };
              }, normalizedSelector, action.value);

              if (selectState.currentVal !== null && selectState.currentVal.toLowerCase() === action.value.toLowerCase()) {
                console.log(`[AI-Step] [${phase}] Skipping select "${normalizedSelector}" — already "${selectState.currentVal}"`);
                if (action.waitAfter && action.waitAfter > 0) await this.sleep(action.waitAfter);
                continue;
              }

              // Guard: if the AI invented an option value that doesn't exist in
              // the <select> and the field already has a valid value, skip to
              // avoid corrupting a good selection (frame.select() with a missing
              // option silently resets the DOM value to "").
              if (!selectState.optionExists) {
                if (selectState.currentVal && selectState.currentVal !== '' && selectState.currentVal.toLowerCase() !== 'select...') {
                  console.log(`[AI-Step] [${phase}] Skipping select "${normalizedSelector}" — option "${action.value}" not found (keeping "${selectState.currentVal}")`);
                } else {
                  console.log(`[AI-Step] [${phase}] Skipping select "${normalizedSelector}" — option "${action.value}" not in DOM`);
                }
                continue;
              }
            } catch (_) {}
          }

          try {
            await this.client.client.evaluate((sel: string) => {
              const el = document.querySelector(sel) as HTMLInputElement;
              if (el && 'value' in el && el.type !== 'checkbox' && el.type !== 'radio') {
                el.value = '';
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
              }
            }, normalizedSelector);
          } catch (_) {
            // Non-fatal — proceed with fill even if clear fails
          }
          await this.client.fillOutField(normalizedSelector, action.value);

          // Post-fill verification for selects: confirm the DOM value stuck.
          if (action.inputType === 'select') {
            try {
              const domVal = await this.client.client.evaluate((sel: string) => {
                const el = document.querySelector(sel) as HTMLSelectElement | null;
                return el ? el.value : '__NOT_FOUND__';
              }, normalizedSelector);
              console.log(`[AI-Step] [${phase}] select post-fill: "${normalizedSelector}" → DOM value="${domVal}"`);
            } catch (_) {}
          }
        }

        // Honour any per-action pause requested by the AI (e.g. waiting for
        // a JS-driven form section to render after a dropdown selection).
        if (action.waitAfter && action.waitAfter > 0) {
          await this.sleep(action.waitAfter);
        }
      } catch (actionErr) {
        if (action.inputType === 'click' && !isRevealPhase) {
          // submitFormByClickingButton can throw even when the page navigated successfully
          // (e.g. "Submit button still there" race during navigation). Don't re-throw —
          // let detectSuccess() determine the real outcome from page state.
          console.log(`[AI-Step] Click action threw (may be a navigation race) — selector="${action.selector}": ${actionErr.toString()}`);
        } else if (isRevealPhase) {
          // Revelation errors are non-fatal — log, record the failure, and continue.
          anyFailed = true;
          console.log(`[AI-Step] Reveal action failed (non-fatal) — type=${action.inputType} selector="${action.selector}": ${actionErr.toString()}`);
        } else {
          // Fill errors are also non-fatal: log and continue with remaining actions.
          // This allows fillProgressiveFields() to run after the main fill pass and
          // pick up any fields that weren't in the DOM yet (e.g. conditional/cascading
          // fields that only appear after a prior selection is made). If the field
          // remains unfilled, the submit will fail and the retry mechanism catches it.
          anyFailed = true;
          console.log(`[AI-Step] Fill action failed (continuing) — selector="${action.selector}" inputType=${action.inputType}: ${actionErr.toString()}`);
        }
      }
    }
    return anyFailed;
  }

  /**
   * Handles custom (non-native) dropdown components by:
   *  1. Scrolling the trigger into view.
   *  2. Clicking the trigger to open the dropdown.
   *  3. Waiting 1 second for the options to render.
   *  4. Finding the first visible element whose text exactly matches `optionText`
   *     and clicking it.
   *
   * The text search is tiered:
   *  - First looks inside the trigger element's subtree (handles inline dropdowns).
   *  - Then searches the full page for standard list/ARIA option roles.
   *  - Finally falls back to any leaf element with matching text.
   */
  private async selectCustomDropdown(triggerSelector: string, optionText: string): Promise<void> {
    // Scroll trigger into view (may be below fold).
    try {
      await this.client.client.evaluate((sel: string) => {
        const el = document.querySelector(sel);
        if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }, triggerSelector);
      await this.sleep(500);
    } catch (_) {}

    // Click the trigger using JS to avoid Puppeteer visibility restrictions.
    await this.client.client.evaluate((sel: string) => {
      const el = document.querySelector(sel);
      if (!el) throw new Error(`Custom dropdown trigger not found: ${sel}`);
      (el as HTMLElement).click();
    }, triggerSelector);

    // Wait for options to render.
    await this.sleep(1000);

    // Find and click the matching option.
    const clicked = await this.client.client.evaluate((sel: string, text: string) => {
      const isVisible = (el: Element) => !!(el as HTMLElement).offsetParent;

      // Tier 1: search within the trigger's own subtree (inline dropdown panels).
      const container = document.querySelector(sel);
      if (container) {
        const inContainer = Array.from(container.querySelectorAll('*'));
        const t1 = inContainer.find(el => el.children.length === 0 && el.textContent?.trim() === text && isVisible(el));
        if (t1) { (t1 as HTMLElement).click(); return true; }
      }

      // Tier 2: standard ARIA / list roles that most dropdown libraries use.
      const ariaTargets = Array.from(document.querySelectorAll(
        'li, [role="option"], [role="menuitem"], [role="listitem"], [role="treeitem"]',
      ));
      const t2 = ariaTargets.find(el => el.textContent?.trim() === text && isVisible(el));
      if (t2) { (t2 as HTMLElement).click(); return true; }

      // Tier 3: any visible leaf element with exactly matching text (broad fallback).
      const allLeaves = Array.from(document.querySelectorAll('*')).filter(
        el => el.children.length === 0 && el.textContent?.trim() === text && isVisible(el),
      );
      if (allLeaves.length > 0) { (allLeaves[0] as HTMLElement).click(); return true; }

      return false;
    }, triggerSelector, optionText) as boolean;

    if (!clicked) {
      throw new Error(`Could not find visible option with text "${optionText}" after clicking ${triggerSelector}`);
    }
  }


  /**
   * After the main fill actions run, some forms reveal additional required fields
   * only after certain selections are made (progressive/conditional disclosure).
   * This method waits briefly, then re-snapshots the form and fills any newly-
   * visible required fields that are still empty — up to MAX_PROGRESSIVE_ROUNDS.
   *
   * Only non-submit (non-click) fill actions are executed so the submit step
   * remains under the caller's control.
   */
  private async fillProgressiveFields(
    aiHelper: AiFormFill,
    messages: any[],
    fieldOverrides: Record<string, string>,
    url: string,
  ): Promise<void> {
    const MAX_PROGRESSIVE_ROUNDS = 2;
    for (let round = 0; round < MAX_PROGRESSIVE_ROUNDS; round++) {
      await this.sleep(2000); // wait for conditional fields to appear

      const { empty: hasEmpty } = await this.hasUnfilledRequiredFields();
      if (!hasEmpty) break;

      // Programmatically fill any empty required <select> elements whose IDs/names
      // we can read directly from the DOM — no AI needed for native selects.
      // This handles fields like "#inquiryType" that the AI tends to hallucinate
      // wrong selectors for (e.g. generating "#cFMInquiryType" by analogy).
      await this.fillEmptyRequiredSelects(fieldOverrides);

      // Similarly fill empty required text/email/tel/textarea inputs using
      // dummy values inferred from the field's label, name, or id — this catches
      // fields the AI either skips or targets with wrong selectors.
      await this.fillEmptyRequiredInputs(fieldOverrides);

      // Re-check after programmatic fills; if everything is now filled, skip AI.
      const { empty: stillEmpty } = await this.hasUnfilledRequiredFields();
      if (!stillEmpty) break;

      const progressScreenshot = await this.safeCaptureBase64();

      // Use the same surgical extraction as Step 6: selects outside any <form>
      // (with their label/wrapper context) + the LAST VISIBLE <form> element.
      // Using only the last visible form avoids the accumulation of stale Marketo
      // form instances that pile up each time the country selector fires its cascade.
      let progressHtml = '';
      try {
        progressHtml = await this.client.client.evaluate(() => {
          function stripStyles(el: Element): Element {
            const clone = el.cloneNode(true) as Element;
            clone.querySelectorAll('style').forEach(s => s.remove());
            return clone;
          }
          function insideForm(el: Element): boolean {
            let p: Element | null = el.parentElement;
            while (p) {
              if (p.tagName === 'FORM') return true;
              p = p.parentElement;
            }
            return false;
          }
          const parts: string[] = [];
          // Selects outside any form (pre-form section)
          Array.from(document.querySelectorAll('select'))
            .filter(sel => !insideForm(sel))
            .forEach(sel => {
              let container: Element = sel;
              for (let i = 0; i < 3; i++) {
                if (!container.parentElement || container.parentElement === document.body) break;
                container = container.parentElement;
              }
              parts.push(stripStyles(container).outerHTML);
            });
          // Only include the last VISIBLE form — earlier instances are stale copies
          // left by Marketo when the country selector triggered a form reload.
          const allForms = Array.from(document.querySelectorAll('form'));
          const visibleForms = allForms.filter(f => !!(f as HTMLElement).offsetParent);
          const targetForm = visibleForms.length > 0
            ? visibleForms[visibleForms.length - 1]
            : allForms[allForms.length - 1];
          if (targetForm) parts.push(stripStyles(targetForm).outerHTML);
          return parts.length > 0 ? parts.join('\n') : document.body.innerHTML;
        }) as string;
      } catch (_) {
        progressHtml = await this.captureFullPageHtml();
      }

      // Diagnostic: log every <select> on the page with its id/name/label
      // so we can verify the AI receives the correct selector.
      try {
        const selectInfo = await this.client.client.evaluate(() =>
          Array.from(document.querySelectorAll('select')).map((sel) => {
            const id = (sel as HTMLSelectElement).id;
            const name = (sel as HTMLSelectElement).name;
            const labelEl = id ? document.querySelector(`label[for="${id}"]`) : null;
            const label = (labelEl as HTMLElement | null)?.innerText?.trim() || '';
            const opts = Array.from((sel as HTMLSelectElement).options)
              .slice(0, 4).map(o => o.text);
            return { id, name, label, opts };
          }),
        );
        console.log(`[AI-Step] Progressive scan selects: ${JSON.stringify(selectInfo)}`);
      } catch (_) {}

      console.log(`[AI-Step] Progressive scan (round ${round + 1}): unfilled required fields detected — html length=${progressHtml.length} (first 800 chars): ${progressHtml.substring(0, 800)}`);
      console.log(`[AI-Step] Progressive scan (round ${round + 1}): calling AI for ${url}`);

      let extraActions: AiFillAction[] = [];
      try {
        const extra = await aiHelper.getFillActions(progressScreenshot, progressHtml, fieldOverrides, messages);
        messages = extra.messages;
        extraActions = extra.actions.filter(a => a.inputType !== 'click');
      } catch (e) {
        console.log(`[AI-Step] Progressive scan AI call failed (non-fatal): ${e}`);
        break;
      }

      if (extraActions.length === 0) break;

      console.log(`[AI-Step] Progressive scan: filling ${extraActions.length} additional field(s)`);
      await this.executeFillActions(extraActions);

      // Run DOM-based fills again after the AI round: the AI may have tried to
      // set selects/inputs with invalid option values (silently resetting them
      // to ""), or may have revealed new required fields.  A second pass repairs
      // any corruption and fills newly-visible fields without another AI call.
      await this.fillEmptyRequiredSelects(fieldOverrides);
      await this.fillEmptyRequiredInputs(fieldOverrides);
    }
  }

  /**
   * Returns true if there are any visible required form fields that still
   * have an empty or placeholder value. Used to detect progressive/conditional
   * fields that appeared after the initial fill actions ran.
   */
  /**
   * Programmatically fill any visible required <select> elements whose current
   * value is empty or a placeholder.  We enumerate these directly from the DOM
   * so we don't rely on the AI to produce correct selectors (which it tends to
   * hallucinate by analogy — e.g. generating "#cFMInquiryType" when the real
   * selector is "#inquiryType").
   *
   * Each empty select is filled with the fieldOverrides value (if present) or
   * the first real option value.  A 1.5 s pause after each fill gives the page
   * JS time to react and reveal dependent fields.
   */
  private async fillEmptyRequiredSelects(fieldOverrides: Record<string, string>): Promise<void> {
    try {
      const emptySelects = await this.client.client.evaluate(() => {
        const PLACEHOLDER_VALUES = ['', 'select...'];
        const results: Array<{ id: string; name: string; selector: string; options: string[] }> = [];

        const selects = Array.from(document.querySelectorAll(
          'select[required], select[aria-required="true"], select.mktoRequired, select.mktoField',
        )) as HTMLSelectElement[];

        for (const sel of selects) {
          if (!sel.offsetParent) continue; // hidden/detached
          const val = sel.value.trim().toLowerCase();
          if (!PLACEHOLDER_VALUES.includes(val)) continue; // already filled

          const options = Array.from(sel.options)
            .map(o => o.value)
            .filter(v => v && v.trim().toLowerCase() !== 'select...');

          if (options.length === 0) continue;

          const id = sel.id;
          const name = sel.name;
          const selector = id ? `#${id}` : (name ? `[name="${name}"]` : '');
          if (!selector) continue;

          results.push({ id, name, selector, options });
        }
        return results;
      }) as Array<{ id: string; name: string; selector: string; options: string[] }>;

      if (emptySelects.length === 0) return;

      for (const field of emptySelects) {
        // Check idempotency (the field might have been filled by a concurrent action)
        try {
          const currentVal = await this.client.client.evaluate((sel: string) => {
            const el = document.querySelector(sel) as HTMLSelectElement | null;
            return el ? el.value : null;
          }, field.selector);
          if (currentVal !== null && currentVal !== '' && currentVal.toLowerCase() !== 'select...') continue;
        } catch (_) {}

        // Prefer fieldOverrides value, then first available option
        const overrideValue = fieldOverrides[field.id] || fieldOverrides[field.name] || fieldOverrides[field.selector];
        const valueToSet = overrideValue || field.options[0];

        console.log(`[AI-Step] Auto-filling empty select "${field.selector}" → "${valueToSet.substring(0, 60)}"`);
        try {
          await this.client.fillOutField(field.selector, valueToSet);

          const domVal = await this.client.client.evaluate((sel: string) => {
            const el = document.querySelector(sel) as HTMLSelectElement | null;
            return el ? el.value : '__NOT_FOUND__';
          }, field.selector);
          console.log(`[AI-Step] Auto-fill select post-fill: "${field.selector}" → DOM value="${domVal}"`);

          // Give the page JS time to react and reveal any dependent fields
          await this.sleep(1500);
        } catch (e) {
          console.log(`[AI-Step] Auto-fill select failed for "${field.selector}" (non-fatal): ${e}`);
        }
      }
    } catch (e) {
      console.log(`[AI-Step] fillEmptyRequiredSelects error (non-fatal): ${e}`);
    }
  }

  /**
   * Programmatically fill visible required text/email/tel/number/textarea inputs
   * that are still empty after the AI fill round.  Dummy values are inferred from
   * the field's label text, name, and id so the fill is contextually appropriate.
   * fieldOverrides values take priority when the field id or name matches a key.
   */
  private async fillEmptyRequiredInputs(fieldOverrides: Record<string, string>): Promise<void> {
    try {
      const emptyInputs = await this.client.client.evaluate(() => {
        const TEXT_TYPES = new Set(['text', 'email', 'tel', 'number', 'url', 'search', 'textarea', '']);

        function getLabelText(el: HTMLElement): string {
          // 1. <label for="id">
          const id = el.id;
          if (id) {
            const label = document.querySelector(`label[for="${id}"]`) as HTMLElement | null;
            if (label) return label.innerText || label.textContent || '';
          }
          // 2. Wrapping <label>
          let p: Element | null = el.parentElement;
          while (p && p !== document.body) {
            if (p.tagName === 'LABEL') return (p as HTMLElement).innerText || '';
            // Marketo wraps fields in .mktoFieldWrap, label is a sibling
            const siblingLabel = p.querySelector('label');
            if (siblingLabel) return (siblingLabel as HTMLElement).innerText || '';
            p = p.parentElement;
          }
          return '';
        }

        const results: Array<{ id: string; name: string; selector: string; labelText: string; inputType: string }> = [];
        const SELECTORS = [
          'input[required]', 'input[aria-required="true"]', 'input.mktoRequired', 'input.mktoField',
          'textarea[required]', 'textarea[aria-required="true"]', 'textarea.mktoRequired',
        ].join(', ');

        const inputs = Array.from(document.querySelectorAll(SELECTORS)) as HTMLInputElement[];
        for (const el of inputs) {
          if (!el.offsetParent) continue; // hidden
          const type = (el.type || '').toLowerCase();
          if (!TEXT_TYPES.has(type)) continue; // skip checkbox/radio/hidden etc.
          if ((el.value || '').trim() !== '') continue; // already filled

          const id = el.id;
          const name = el.name || '';
          const selector = id ? `#${id}` : (name ? `[name="${name}"]` : '');
          if (!selector) continue;

          results.push({ id, name, selector, labelText: getLabelText(el), inputType: el.tagName.toLowerCase() });
        }
        return results;
      }) as Array<{ id: string; name: string; selector: string; labelText: string; inputType: string }>;

      if (emptyInputs.length === 0) return;

      for (const field of emptyInputs) {
        // Idempotency: skip if already filled
        try {
          const currentVal = await this.client.client.evaluate((sel: string) => {
            const el = document.querySelector(sel) as HTMLInputElement | null;
            return el ? el.value : null;
          }, field.selector);
          if (currentVal !== null && currentVal.trim() !== '') continue;
        } catch (_) {}

        // fieldOverrides by id or name take priority
        const overrideValue = fieldOverrides[field.id] || fieldOverrides[field.name] || fieldOverrides[field.selector];
        const valueToSet = overrideValue || this.inferDummyValue(field.id, field.name, field.labelText, field.inputType);

        console.log(`[AI-Step] Auto-filling empty input "${field.selector}" (label: "${field.labelText.trim().substring(0, 40)}") → "${valueToSet.substring(0, 60)}"`);
        try {
          await this.client.fillOutField(field.selector, valueToSet);
        } catch (e) {
          console.log(`[AI-Step] Auto-fill input failed for "${field.selector}" (non-fatal): ${e}`);
        }
      }
    } catch (e) {
      console.log(`[AI-Step] fillEmptyRequiredInputs error (non-fatal): ${e}`);
    }
  }

  /** Infer an appropriate dummy value from the field's label/name/id. */
  private inferDummyValue(id: string, name: string, labelText: string, inputType: string): string {
    const hint = `${id} ${name} ${labelText}`.toLowerCase();

    if (/email/.test(hint)) return 'testrun@example.com';
    if (/phone|tel|mobile|cell|fax/.test(hint)) return '555-555-0192';
    if (/zip|postal|post.?code/.test(hint)) return '10001';
    if (/city|town/.test(hint)) return 'New York';
    if (/state|province|region/.test(hint) && inputType !== 'select') return 'NY';
    if (/address|street|addr/.test(hint)) return '123 Main Street';
    if (/company|organization|org|institution|firm|business|employer|account/.test(hint)) return 'Test Company';
    if (/first.?name|given.?name|fname/.test(hint)) return 'Test';
    if (/last.?name|family.?name|surname|lname/.test(hint)) return 'Run';
    if (/full.?name|your.?name/.test(hint)) return 'Test Run';
    if (/name/.test(hint)) return 'Test Run';
    if (/title|subject/.test(hint)) return 'Test Subject';
    if (/product|item|model|part/.test(hint)) return 'Test Product';
    if (/serial|account.?num|customer.?num|id/.test(hint)) return '12345';
    if (/url|website|web.?site/.test(hint)) return 'https://example.com';
    if (/message|comment|feedback|description|details|notes|inquiry|question|issue/.test(hint)) {
      return 'This is a test inquiry.';
    }
    if (inputType === 'textarea') return 'This is a test message.';
    return 'Test Value';
  }

  private async hasUnfilledRequiredFields(): Promise<{ empty: boolean; details: string }> {
    try {
      const result = await this.client.client.evaluate(() => {
        const PLACEHOLDER_TEXTS = ['select...', 'please select', '-- select --', '- select -'];
        const unfilled: string[] = [];

        // Deduplicate by ID — Marketo appends fresh form instances when the
        // country selector changes, leaving old (hidden) copies in the DOM.
        // Counting those copies inflates the count and confuses the AI.
        const seenIds = new Set<string>();

        // Only target actual form controls — Marketo also puts .mktoRequired on
        // wrapper <div>s which have no .value; scanning those produces hundreds
        // of false positives.
        const required = Array.from(document.querySelectorAll(
          'input[required], input[aria-required="true"], input.mktoRequired, input.mktoField,' +
          'select[required], select[aria-required="true"], select.mktoRequired,' +
          'textarea[required], textarea[aria-required="true"], textarea.mktoRequired',
        ));
        for (const el of required) {
          const rect = (el as HTMLElement).getBoundingClientRect();
          // Skip invisible elements (old, hidden form instances have zero size
          // or are scrolled way above the viewport).
          if (rect.width === 0 || rect.height === 0) continue;
          if (!(el as HTMLElement).offsetParent) continue;

          // Deduplicate by element ID
          const elId = (el as HTMLInputElement).id;
          if (elId) {
            if (seenIds.has(elId)) continue;
            seenIds.add(elId);
          }

          const val = ((el as HTMLInputElement).value || '').trim().toLowerCase();
          if (!val || PLACEHOLDER_TEXTS.some(p => val.startsWith(p))) {
            const tag = el.tagName.toLowerCase();
            const id = elId ? `#${elId}` : '';
            const name = (el as HTMLInputElement).name ? `[name="${(el as HTMLInputElement).name}"]` : '';
            unfilled.push(`${tag}${id || name} val="${val}"`);
          }
        }
        // Also check visible custom dropdown triggers still showing a placeholder
        const seenTriggers = new Set<string>();
        const triggers = Array.from(document.querySelectorAll(
          '[class*="dropdown"] [class*="trigger"], [class*="dropdown"] [class*="selected"], ' +
          '[class*="select"] [class*="placeholder"], .mktoField',
        ));
        for (const el of triggers) {
          const rect = (el as HTMLElement).getBoundingClientRect();
          if (rect.width === 0 || rect.height === 0) continue;
          if (!(el as HTMLElement).offsetParent) continue;
          const text = ((el as HTMLElement).innerText || '').trim().toLowerCase();
          if (PLACEHOLDER_TEXTS.some(p => text.startsWith(p))) {
            const key = text.substring(0, 30);
            if (seenTriggers.has(key)) continue;
            seenTriggers.add(key);
            unfilled.push(`custom-trigger "${(el as HTMLElement).innerText?.trim().substring(0, 40)}"`);
          }
        }
        return unfilled;
      }) as string[];

      if (result.length > 0) {
        console.log(`[AI-Step] Unfilled required fields (${result.length}): ${result.join(' | ')}`);
        return { empty: true, details: result.join(', ') };
      }
      return { empty: false, details: '' };
    } catch (_) {
      return { empty: false, details: '' };
    }
  }

  private async detectSuccess(preSubmitUrl: string): Promise<boolean> {
    try {
      const currentUrl = await this.client.client.url();
      // URL navigation is the strongest signal — the form submitted and redirected.
      if (currentUrl !== preSubmitUrl) return true;

      // A visible success message is a reliable in-page signal.
      const hasSuccessMessage: boolean = await this.client.client.evaluate(() => {
        const text = (document.body as HTMLElement).innerText.toLowerCase();
        return ['thank you', 'thanks!', 'success', 'submitted', 'received', 'confirmed']
          .some((p) => text.includes(p));
      }) as boolean;
      if (hasSuccessMessage) return true;

      // NOTE: We intentionally do NOT use "formGone" (no <form> elements) as a
      // standalone success signal. Many modern pages use custom form components
      // that never render a <form> element, so formGone would always be true and
      // produce false positives on every attempt.
      return false;
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

        // Ignore text that comes from interactive elements (button labels,
        // select option lists, etc.) — these cause false positives like "Submit"
        // being treated as a validation error.
        function isInteractive(el: Element): boolean {
          let node: Element | null = el;
          while (node && node !== document.body) {
            const tag = node.tagName.toLowerCase();
            if (['button', 'select', 'input', 'a', 'label'].includes(tag)) return true;
            if ((node as HTMLElement).getAttribute('role') === 'option') return true;
            if ((node as HTMLElement).getAttribute('type') === 'submit') return true;
            node = node.parentElement;
          }
          return false;
        }

        const found: string[] = [];
        for (const sel of selectors) {
          document.querySelectorAll(sel).forEach((el) => {
            if (isInteractive(el)) return;
            const t = ((el as HTMLElement).innerText || '').trim();
            // Also skip if the text is just a single word that could be a button label
            if (t && t.length > 2 && !/^(submit|next|back|send|ok|yes|no|cancel)$/i.test(t)) {
              found.push(t);
            }
          });
        }
        return found.slice(0, 5).join('; ');
      }) as string;
      return text;
    } catch (_) {
      return '';
    }
  }

  /**
   * Attempts to dismiss cookie/privacy consent banners by clicking common
   * "Accept All" / "Accept" buttons that various consent management platforms
   * (OneTrust, Cookiebot, TrustArc, etc.) place on enterprise websites.
   * Runs silently — any failure is non-fatal.
   */
  private async dismissCookieBanners(): Promise<void> {
    const ACCEPT_SELECTORS = [
      '#onetrust-accept-btn-handler',         // OneTrust (most common on enterprise/pharma sites)
      '#accept-recommended-btn-handler',       // OneTrust alternate
      '.js-accept-recommended-btn-handler',    // OneTrust class-based
      '#CybotCookiebotDialogBodyLevelButtonLevelOptinAllowAll', // Cookiebot
      '#cookie-accept',
      '#cookieAccept',
      '.js-cookie-accept',
      'button[data-consent-action="accept"]',
      'button[data-action="accept-all"]',
    ];
    for (const sel of ACCEPT_SELECTORS) {
      try {
        const clicked = await this.client.client.evaluate((s: string) => {
          const el = document.querySelector(s) as HTMLElement | null;
          if (el) { el.click(); return true; }
          return false;
        }, sel);
        if (clicked) {
          console.log(`[AI-Step] Dismissed cookie banner via ${sel}`);
          await this.sleep(800);
          return;
        }
      } catch (_) {}
    }
  }

  /**
   * Captures page HTML focused on the main content area, not the full body.
   *
   * Strategy (in order):
   *  1. <main> or [role="main"] — excludes header/nav entirely.
   *  2. The element at the center of the current viewport (works at any scroll
   *     position — at 75% scroll this lands in the form section). We walk up
   *     to the nearest section/article/div that's taller than half the viewport
   *     so we get a meaningful container, not just a leaf node.
   *  3. Full document.body fallback.
   *
   * Same-origin iframe contents are appended in all cases.
   */
  private async captureFullPageHtml(): Promise<string> {
    return (await this.client.client.evaluate(() => {
      function appendIframes(base: string): string {
        const iframes = Array.from(document.querySelectorAll('iframe'));
        for (const iframe of iframes) {
          try {
            const doc = (iframe as HTMLIFrameElement).contentDocument;
            if (doc && doc.body) {
              const src = iframe.getAttribute('src') || iframe.id || 'unnamed';
              base += `\n<!-- IFRAME[${src}] -->\n${doc.body.innerHTML}`;
            }
          } catch (_) { /* cross-origin — skip */ }
        }
        return base;
      }

      // Strategy 1: semantic main content element
      const mainEl = document.querySelector('main, [role="main"], #main-content, #content, .main-content');
      if (mainEl) {
        return appendIframes((mainEl as HTMLElement).innerHTML);
      }

      // Strategy 2: element at center of the current viewport
      const centerEl = document.elementFromPoint(window.innerWidth / 2, window.innerHeight / 2);
      if (centerEl && centerEl !== document.body && centerEl !== document.documentElement) {
        // Walk up until we find a section-level container (tall enough to be meaningful)
        let container: Element | null = centerEl;
        while (container && container !== document.body) {
          const tag = container.tagName.toLowerCase();
          const height = (container as HTMLElement).offsetHeight;
          if (['section', 'article', 'main', 'form'].includes(tag) || height > window.innerHeight * 0.4) {
            return appendIframes((container as HTMLElement).innerHTML);
          }
          container = container.parentElement;
        }
      }

      // Strategy 3: full body fallback
      return appendIframes(document.body.innerHTML);
    })) as string;
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
    revealActions: AiFillAction[],
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
      const promotedSteps = this.buildPromotedSteps(url, revealActions, fillActions);
      // The 'exposeOnPass:' prefix tells the cog-mechanism's StepResponse.fromProto
      // to keep this record in the run log even when the step outcome is Passed.
      records.push(this.keyValue('exposeOnPass:promotedSteps', 'Promoted Concrete Steps', {
        stepsJson: JSON.stringify(promotedSteps),
      }));
    }

    return records;
  }

  /**
   * Builds the concrete step definitions that replace this AI step when
   * cacheStrategy === "promote". Revelation steps (click/select/scroll/focusFrame)
   * are prepended before the fill steps so the promoted scenario is fully
   * self-contained.
   *
   * waitAfter on a reveal action is applied as `waitFor` (seconds) on the
   * immediately following promoted step.
   */
  private buildPromotedSteps(url: string, revealActions: AiFillAction[], fillActions: AiFillAction[]): any[] {
    const steps: any[] = [
      {
        stepId: 'NavigateToPage',
        name: 'Navigate to a webpage',
        cog: 'automatoninc/web',
        data: { webPageUrl: url },
      },
    ];

    // pendingWaitFor carries a wait time (in seconds) from the PREVIOUS action's waitAfter.
    // It is applied to the NEXT promoted step, so that the scenario pauses before
    // executing that step while JS-driven content (triggered by the previous action) loads.
    let pendingWaitFor: number | undefined;

    const buildStep = (stepObj: any): any => {
      if (pendingWaitFor) {
        stepObj.waitFor = pendingWaitFor;
        pendingWaitFor = undefined;
      }
      return stepObj;
    };

    // ── Revelation steps ───────────────────────────────────────────────────────
    for (const action of revealActions) {
      if (action.inputType === 'focusFrame') {
        steps.push(buildStep({
          stepId: 'FocusOnFrame',
          name: 'Focus on frame',
          cog: 'automatoninc/web',
          data: { domQuerySelector: action.selector },
        }));
      } else if (action.inputType === 'click') {
        steps.push(buildStep({
          stepId: 'ClickOnElement',
          name: 'Click an element on a page',
          cog: 'automatoninc/web',
          data: { domQuerySelector: action.selector },
        }));
      } else if (action.inputType === 'selectCustomDropdown') {
        // Promoted as a ClickOnElement (trigger) — the option selection is done inline
        // by the custom dropdown logic and cannot be expressed as a single concrete step.
        steps.push(buildStep({
          stepId: 'ClickOnElement',
          name: 'Click an element on a page',
          cog: 'automatoninc/web',
          data: { domQuerySelector: action.selector },
        }));
      } else if (action.inputType === 'select') {
        steps.push(buildStep({
          stepId: 'EnterValueIntoField',
          name: 'Fill out a form field',
          cog: 'automatoninc/web',
          data: { domQuerySelector: action.selector, value: action.value },
        }));
      } else if (action.inputType === 'scroll') {
        // Map to ScrollTo at 75% — a reasonable approximation for "scroll to reveal"
        // since the target element is typically in the lower portion of the page.
        steps.push(buildStep({
          stepId: 'ScrollTo',
          name: 'Scroll on a web page',
          cog: 'automatoninc/web',
          data: { depth: 75, units: '%' },
        }));
      }

      // Set pendingWaitFor AFTER pushing this step so it appears on the NEXT step,
      // giving the page time to react (e.g. form appearing after a dropdown selection).
      if (action.waitAfter) {
        pendingWaitFor = Math.ceil(action.waitAfter / 1000);
      }
    }

    // ── Fill steps ─────────────────────────────────────────────────────────────
    for (const action of fillActions) {
      if (action.inputType === 'click') {
        steps.push(buildStep({
          stepId: 'SubmitFormByClickingButton',
          name: 'Submit a form by clicking a button',
          cog: 'automatoninc/web',
          data: { domQuerySelector: action.selector },
        }));
      } else if (action.inputType === 'focusFrame') {
        steps.push(buildStep({
          stepId: 'FocusOnFrame',
          name: 'Focus on frame',
          cog: 'automatoninc/web',
          data: { domQuerySelector: action.selector },
        }));
      } else if (action.inputType === 'selectCustomDropdown') {
        steps.push(buildStep({
          stepId: 'ClickOnElement',
          name: 'Click an element on a page',
          cog: 'automatoninc/web',
          data: { domQuerySelector: action.selector },
        }));
      } else {
        steps.push(buildStep({
          stepId: 'EnterValueIntoField',
          name: 'Fill out a form field',
          cog: 'automatoninc/web',
          data: { domQuerySelector: action.selector, value: action.value },
        }));
      }
    }

    return steps;
  }
}

export { NavigateAndSubmitFormWithAI as Step };
