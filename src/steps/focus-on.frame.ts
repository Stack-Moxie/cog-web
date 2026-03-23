import { BaseStep, Field, StepInterface } from '../core/base-step';
import { Step, RunStepResponse, FieldDefinition, StepDefinition, StepRecord } from '../proto/cog_pb';

export class FocusOnFrame extends BaseStep implements StepInterface {

  protected stepName: string = 'Focus on frame';
  // tslint:disable-next-line:max-line-length
  protected stepExpression: string = 'focus on the (?<domQuerySelector>.+) frame';
  protected stepType: StepDefinition.Type = StepDefinition.Type.ACTION;
  protected actionList: string[] = ['interact'];
  protected targetObject: string = 'Focus on frame';
  protected expectedFields: Field[] = [{
    field: 'domQuerySelector',
    type: FieldDefinition.Type.STRING,
    description: 'The iframe\'s DOM query selector, or "main" for the main frame',
  }];

  async executeStep(step: Step): Promise<RunStepResponse> {
    const stepData: any = step.getData().toJavaScript();
    const iframeSelector: string = stepData.domQuerySelector;

    try {
      await this.client.focusFrame(iframeSelector);
      let binaryRecord;
      try {
        const screenshot = await this.client.safeScreenshot({ type: 'jpeg', encoding: 'binary', quality: 60 });
        binaryRecord = this.binary('screenshot', 'Screenshot', 'image/jpeg', screenshot);
      } catch (_) {}
      const record = this.createRecord(iframeSelector);
      const orderedRecord = this.createOrderedRecord(iframeSelector, stepData['__stepOrder']);
      return this.pass('Successfully focused on frame %s', [iframeSelector], binaryRecord ? [binaryRecord, record, orderedRecord] : [record, orderedRecord]);
    } catch (e) {
      let binaryRecord;
      try {
        const screenshot = await this.client.safeScreenshot({ type: 'jpeg', encoding: 'binary', quality: 60 });
        binaryRecord = this.binary('screenshot', 'Screenshot', 'image/jpeg', screenshot);
      } catch (_) {}
      return this.error(
        'Unable to focus on frame %s, the frame may no longer exist on the page: %s',
        [iframeSelector, e.toString()],
        binaryRecord ? [binaryRecord] : [],
      );
    }
  }

  public createRecord(frame): StepRecord {
    const obj = {
      frame,
    };
    const record = this.keyValue('form', 'Focused on Frame', obj);

    return record;
  }

  public createOrderedRecord(frame, stepOrder = 1): StepRecord {
    const obj = {
      frame,
    };
    const record = this.keyValue(`form.${stepOrder}`, `Focused on Frame from Step ${stepOrder}`, obj);

    return record;
  }

}

export { FocusOnFrame as Step };
