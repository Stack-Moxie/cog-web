import { BaseStep, ExpectedRecord, Field, StepInterface } from '../core/base-step';
import { Step, RunStepResponse, FieldDefinition, StepDefinition, StepRecord, RecordDefinition } from '../proto/cog_pb';

export class NavigateToPage extends BaseStep implements StepInterface {

  protected stepName: string = 'Navigate to a webpage';
  // tslint:disable-next-line:max-line-length
  protected stepExpression: string = 'navigate to (?<webPageUrl>.+)';
  protected stepType: StepDefinition.Type = StepDefinition.Type.ACTION;
  protected actionList: string[] = ['navigate'];
  protected targetObject: string = 'Navigate to page';
  protected expectedFields: Field[] = [{
    field: 'webPageUrl',
    type: FieldDefinition.Type.URL,
    description: 'Page URL',
  }];

  protected expectedRecords: ExpectedRecord[] = [{
    id: 'form',
    type: RecordDefinition.Type.KEYVALUE,
    fields: [{
      field: 'url',
      type: FieldDefinition.Type.STRING,
      description: 'Url to navigate to',
    }],
    dynamicFields: true,
  }];

  async executeStep(step: Step): Promise<RunStepResponse> {
    const stepData: any = step.getData().toJavaScript();
    const url: string = stepData.webPageUrl;
    const throttle: boolean = stepData.throttle || false;
    const maxInflightRequests: number = stepData.maxInflightRequests || 0;
    const passOn404: boolean = stepData.passOn404 || false;

    // Navigate to URL.
    try {
      console.time('time');
      console.log('>>>>> STARTED TIMER FOR NAVIGATE-TO-PAGE STEP');
      await this.client.navigateToUrl(url, throttle, maxInflightRequests);

      // Stop any streaming media before taking the screenshot. A streaming video
      // (e.g. a 7.2 Mbps homepage hero video) keeps Chrome's rendering pipeline
      // busy and causes Page.captureScreenshot to hang indefinitely.
      // Use a per-call timeout so a busy Chrome fails fast (< protocolTimeout).
      try {
        await Promise.race([
          this.client.client.evaluate(() => {
            document.querySelectorAll<HTMLMediaElement>('video, audio').forEach(el => {
              try { el.pause(); el.src = ''; el.load(); } catch (_) {}
            });
          }),
          new Promise<never>((_, reject) => setTimeout(() => reject(new Error('per-call timeout')), 5000)),
        ]);
        console.log('>>>>> checkpoint 5.5: stopped streaming media');
      } catch (stopMediaError) {
        console.log('>>>>> checkpoint 5.5: could not stop media (non-fatal):', stopMediaError.message);
      }

      // Take screenshot. Some pages (e.g. those with streaming video) can cause
      // Page.captureScreenshot to hang indefinitely. We treat a screenshot failure
      // as non-fatal so a successful navigation still passes.
      // Use a per-call timeout so a busy Chrome fails fast (< protocolTimeout).
      let binaryRecord;
      try {
        const screenshot = await Promise.race([
          this.client.client.screenshot({ type: 'jpeg', encoding: 'binary', quality: 60 }),
          new Promise<never>((_, reject) => setTimeout(() => reject(new Error('per-call timeout')), 8000)),
        ]);
        binaryRecord = this.binary('screenshot', 'Screenshot', 'image/jpeg', screenshot);
        console.log('>>>>> checkpoint 6: finished taking screenshot and making binary record');
      } catch (screenshotError) {
        console.log('>>>>> checkpoint 6: screenshot failed (non-fatal):', screenshotError.message);
      }
      console.timeLog('time');

      const lastResponse = this.client.client['___lastResponse'];
      const status = lastResponse ? await lastResponse.status() : 200;
      console.log('>>>>> checkpoint 7: finished getting status, ending timer');
      console.timeEnd('time');
      if (status === 404 && !passOn404) {
        return this.fail('%s returned an Error: 404 Not Found', [url], binaryRecord ? [binaryRecord] : []);
      }
      const record = this.createRecord(url);
      const orderedRecord = this.createOrderedRecord(url, stepData['__stepOrder']);
      return this.pass('Successfully navigated to %s', [url], binaryRecord ? [binaryRecord, record, orderedRecord] : [record, orderedRecord]);
    } catch (e) {
      try {
        const screenshot = await this.client.client.screenshot({ type: 'jpeg', encoding: 'binary', quality: 60 });
        const binaryRecord = this.binary('screenshot', 'Screenshot', 'image/jpeg', screenshot);
        return this.error(
          'There was a problem navigating to %s: %s',
          [url, e.toString()],
          [binaryRecord],
        );
      } catch (screenshotError) {
        return this.error(
          'There was a problem navigating to %s: %s',
          [url, e.toString()],
        );
      }
    }
  }

  public createRecord(url): StepRecord {
    const obj = {
      url,
    };
    const record = this.keyValue('form', 'Navigated to Page', obj);

    return record;
  }

  public createOrderedRecord(url, stepOrder = 1): StepRecord {
    const obj = {
      url,
    };
    const record = this.keyValue(`form.${stepOrder}`, `Navigated to Page from Step ${stepOrder}`, obj);

    return record;
  }

}

export { NavigateToPage as Step };
