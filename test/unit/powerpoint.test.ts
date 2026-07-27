import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";
import { insertText } from "../../src/taskpane/powerpoint";

/* global describe, global, it */

type MockTextShape = {
  text: string;
};

const shapes: MockTextShape[] = [];
const selectedSlide = {
  shapes: {
    addTextBox: function (text: string) {
      const shape: MockTextShape = { text };
      shapes.push(shape);
    },
    items: shapes,
  },
};

const powerpointMockContext = {
  presentation: {
    getSelectedSlides: function () {
      return {
        getItemAt: function () {
          return selectedSlide;
        },
      };
    },
  },
  slides: {
    items: [selectedSlide],
  },
};

const PowerPointMockData = {
  context: powerpointMockContext,
  onReady: async function () {},
  run: async function (callback: (context: typeof powerpointMockContext) => Promise<void> | void) {
    await callback(this.context);
  },
};

describe(`PowerPoint`, function () {
  it("Inserts text", async function () {
    const officeMock = new OfficeMockObject(PowerPointMockData);
    global.PowerPoint = officeMock as any;

    await insertText("Hello PowerPoint");

    assert.strictEqual(shapes[0].text, "Hello PowerPoint");
  });
});
