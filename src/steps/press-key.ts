import { BaseStep, Field, StepInterface } from '../core/base-step';
import { Step, RunStepResponse, FieldDefinition, StepDefinition } from '../proto/cog_pb';

export class PressKey extends BaseStep implements StepInterface {

  protected stepName: string = 'Press a key';
  protected stepExpression: string = 'press the (?<key>.+) key';
  protected stepType: StepDefinition.Type = StepDefinition.Type.ACTION;
  protected actionList: string[] = ['interact'];
  protected targetObject: string = 'Press a key';
  protected expectedFields: Field[] = [{
    field: 'key',
    type: FieldDefinition.Type.STRING,
    description: 'Key to be pressed',
  }];

  async executeStep(step: Step): Promise<RunStepResponse> {
    const stepData: any = step.getData().toJavaScript();
    const key: string = stepData.key;

    try {
      await this.client.pressKey(key);
      let binaryRecord;
      try {
        const screenshot = await this.client.safeScreenshot({ type: 'jpeg', encoding: 'binary', quality: 60 });
        binaryRecord = this.binary('screenshot', 'Screenshot', 'image/jpeg', screenshot);
      } catch (_) {}
      return this.pass('Successfully pressed key: %s', [key], binaryRecord ? [binaryRecord] : []);
    } catch (e) {
      let binaryRecord;
      try {
        const screenshot = await this.client.safeScreenshot({ type: 'jpeg', encoding: 'binary', quality: 60 });
        binaryRecord = this.binary('screenshot', 'Screenshot', 'image/jpeg', screenshot);
      } catch (_) {}
      return this.error(
        'There was a problem pressing key %s: %s',
        [key, e.toString()],
        binaryRecord ? [binaryRecord] : [],
      );
    }
  }

}

export { PressKey as Step };
