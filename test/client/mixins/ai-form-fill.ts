import * as chai from 'chai';
import { default as sinon } from 'ts-sinon';
import * as sinonChai from 'sinon-chai';
import 'mocha';

import { AiFormFill, AiFillAction } from '../../../src/client/mixins/ai-form-fill';

chai.use(sinonChai);

describe('AiFormFill', () => {
  const expect = chai.expect;

  const VALID_ACTIONS: AiFillAction[] = [
    { selector: '#given-name', value: 'Alex', inputType: 'text' },
    { selector: '#email', value: 'alex@example.com', inputType: 'text' },
    { selector: 'button[type="submit"]', value: 'submit', inputType: 'click' },
  ];

  function makeOpenAiResponse(content: string) {
    return {
      choices: [{ message: { content }, finish_reason: 'stop' }],
      usage: { total_tokens: 100, prompt_tokens: 80, completion_tokens: 20 },
    };
  }

  // ── Constructor ───────────────────────────────────────────────────────────────

  describe('constructor', () => {
    afterEach(() => {
      delete process.env.AZURE_OPENAI_ENDPOINT;
      delete process.env.AZURE_OPENAI_API_KEY;
      delete process.env.AZURE_OPENAI_DEPLOYMENT_NAME;
    });

    it('throws when AZURE_OPENAI_ENDPOINT is missing', () => {
      process.env.AZURE_OPENAI_API_KEY = 'key';
      process.env.AZURE_OPENAI_DEPLOYMENT_NAME = 'gpt-4o';
      expect(() => new AiFormFill()).to.throw(/AZURE_OPENAI_ENDPOINT/);
    });

    it('throws when AZURE_OPENAI_API_KEY is missing', () => {
      process.env.AZURE_OPENAI_ENDPOINT = 'https://test.openai.azure.com/';
      process.env.AZURE_OPENAI_DEPLOYMENT_NAME = 'gpt-4o';
      expect(() => new AiFormFill()).to.throw(/AZURE_OPENAI_API_KEY/);
    });

    it('throws when AZURE_OPENAI_DEPLOYMENT_NAME is missing', () => {
      process.env.AZURE_OPENAI_ENDPOINT = 'https://test.openai.azure.com/';
      process.env.AZURE_OPENAI_API_KEY = 'key';
      expect(() => new AiFormFill()).to.throw(/AZURE_OPENAI_DEPLOYMENT_NAME/);
    });

    it('constructs successfully when all env vars are present', () => {
      process.env.AZURE_OPENAI_ENDPOINT = 'https://test.openai.azure.com/';
      process.env.AZURE_OPENAI_API_KEY = 'test-key';
      process.env.AZURE_OPENAI_DEPLOYMENT_NAME = 'gpt-4o';
      expect(() => new AiFormFill()).to.not.throw();
    });
  });

  // ── getFillActions ────────────────────────────────────────────────────────────

  describe('getFillActions', () => {
    let aiFormFill: AiFormFill;
    let createStub: sinon.SinonStub;

    beforeEach(() => {
      process.env.AZURE_OPENAI_ENDPOINT = 'https://test.openai.azure.com/';
      process.env.AZURE_OPENAI_API_KEY = 'test-key';
      process.env.AZURE_OPENAI_DEPLOYMENT_NAME = 'gpt-4o';

      aiFormFill = new AiFormFill();

      // Replace the internal openai client so no HTTP calls are made
      createStub = sinon.stub().resolves(makeOpenAiResponse(JSON.stringify(VALID_ACTIONS)));
      (aiFormFill as any).openai = { chat: { completions: { create: createStub } } };
    });

    afterEach(() => {
      sinon.restore();
      delete process.env.AZURE_OPENAI_ENDPOINT;
      delete process.env.AZURE_OPENAI_API_KEY;
      delete process.env.AZURE_OPENAI_DEPLOYMENT_NAME;
    });

    it('returns parsed fill actions from the AI response', async () => {
      const result = await aiFormFill.getFillActions('screenshot', '<form></form>', {}, []);

      expect(result.actions).to.deep.equal(VALID_ACTIONS);
      expect(result.actions).to.have.length(3);
    });

    it('builds the initial message array (system + user) on the first call', async () => {
      await aiFormFill.getFillActions('screenshot', '<form></form>', {}, []);

      const [callArgs] = createStub.firstCall.args;
      expect(callArgs.messages).to.have.length(2);
      expect(callArgs.messages[0].role).to.equal('system');
      expect(callArgs.messages[1].role).to.equal('user');
    });

    it('appends to existing messages on a retry call (does not rebuild initial messages)', async () => {
      const existingMessages = [
        { role: 'system', content: 'sys' },
        { role: 'user', content: 'first user msg' },
        { role: 'assistant', content: 'first assistant msg' },
      ];

      await aiFormFill.getFillActions('screenshot', '<form></form>', {}, existingMessages);

      const [callArgs] = createStub.firstCall.args;
      // Should still be 3 messages (not rebuilt to 2)
      expect(callArgs.messages).to.have.length(3);
    });

    it('appends the assistant reply to the returned messages array', async () => {
      const result = await aiFormFill.getFillActions('screenshot', '<form></form>', {}, []);

      const lastMsg = result.messages[result.messages.length - 1];
      expect(lastMsg.role).to.equal('assistant');
    });

    it('includes field overrides in the user message text', async () => {
      const overrides = { Country: 'US', State: 'Georgia' };
      await aiFormFill.getFillActions('screenshot', '<form></form>', overrides, []);

      const [callArgs] = createStub.firstCall.args;
      const userMsg = callArgs.messages[1];
      const textPart = userMsg.content.find((c: any) => c.type === 'text');
      expect(textPart.text).to.include('Country');
      expect(textPart.text).to.include('Georgia');
    });

    it('does not include override text when fieldOverrides is empty', async () => {
      await aiFormFill.getFillActions('screenshot', '<form></form>', {}, []);

      const [callArgs] = createStub.firstCall.args;
      const textPart = callArgs.messages[1].content.find((c: any) => c.type === 'text');
      expect(textPart.text).to.not.include('CRITICAL');
    });

    it('includes the screenshot as an image_url content part', async () => {
      await aiFormFill.getFillActions('base64data', '<form></form>', {}, []);

      const [callArgs] = createStub.firstCall.args;
      const imagePart = callArgs.messages[1].content.find((c: any) => c.type === 'image_url');
      expect(imagePart.image_url.url).to.include('base64data');
    });

    it('strips markdown code fences from the AI response before parsing', async () => {
      const fenced = '```json\n' + JSON.stringify(VALID_ACTIONS) + '\n```';
      createStub.resolves(makeOpenAiResponse(fenced));

      const result = await aiFormFill.getFillActions('screenshot', '<form></form>', {}, []);
      expect(result.actions).to.deep.equal(VALID_ACTIONS);
    });

    it('throws when the AI returns a non-array JSON value', async () => {
      createStub.resolves(makeOpenAiResponse('{"error": "bad response"}'));

      await expect(
        aiFormFill.getFillActions('screenshot', '<form></form>', {}, []),
      ).to.be.rejectedWith(/not a JSON array/i);
    });

    it('throws when the AI returns invalid JSON', async () => {
      createStub.resolves(makeOpenAiResponse('this is not json'));

      await expect(
        aiFormFill.getFillActions('screenshot', '<form></form>', {}, []),
      ).to.be.rejected;
    });

    it('truncates form HTML to 50 000 chars before sending to AI', async () => {
      const longHtml = 'x'.repeat(60000);
      await aiFormFill.getFillActions('screenshot', longHtml, {}, []);

      const [callArgs] = createStub.firstCall.args;
      const textPart = callArgs.messages[1].content.find((c: any) => c.type === 'text');
      // The HTML fragment in the text should be capped at 50000 chars
      expect(textPart.text.length).to.be.lessThan(60000);
    });
  });

  // ── buildRetryMessage ─────────────────────────────────────────────────────────

  describe('buildRetryMessage', () => {
    let aiFormFill: AiFormFill;

    before(() => {
      process.env.AZURE_OPENAI_ENDPOINT = 'https://test.openai.azure.com/';
      process.env.AZURE_OPENAI_API_KEY = 'test-key';
      process.env.AZURE_OPENAI_DEPLOYMENT_NAME = 'gpt-4o';
      aiFormFill = new AiFormFill();
    });

    after(() => {
      delete process.env.AZURE_OPENAI_ENDPOINT;
      delete process.env.AZURE_OPENAI_API_KEY;
      delete process.env.AZURE_OPENAI_DEPLOYMENT_NAME;
    });

    it('returns a user-role message', () => {
      const msg = aiFormFill.buildRetryMessage('screenshot', 'Field required');
      expect(msg.role).to.equal('user');
    });

    it('includes the screenshot as an image_url content part', () => {
      const msg = aiFormFill.buildRetryMessage('base64retrydata', '');
      const imagePart = msg.content.find((c: any) => c.type === 'image_url');
      expect(imagePart.image_url.url).to.include('base64retrydata');
    });

    it('includes the error text in the message when provided', () => {
      const msg = aiFormFill.buildRetryMessage('screenshot', 'Email is required');
      const textPart = msg.content.find((c: any) => c.type === 'text');
      expect(textPart.text).to.include('Email is required');
    });

    it('omits the "Visible errors" line when errorText is empty', () => {
      const msg = aiFormFill.buildRetryMessage('screenshot', '');
      const textPart = msg.content.find((c: any) => c.type === 'text');
      expect(textPart.text).to.not.include('Visible errors');
    });
  });
});
