import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";
import { insertText } from "../../src/taskpane/word";

/* global describe, global, it, Word */

type MockParagraph = {
  text: string;
  insertLocation?: string;
};

const wordMockContext = {
  document: {
    body: {
      paragraph: {
        text: "",
      } as MockParagraph,
      insertParagraph: function (paragraphText: string, insertLocation: string): MockParagraph {
        this.paragraph.text = paragraphText;
        this.paragraph.insertLocation = insertLocation;
        return this.paragraph;
      },
    },
  },
};

const WordMockData = {
  context: wordMockContext,
  InsertLocation: {
    end: "End",
  },
  run: async function (callback: (context: typeof wordMockContext) => Promise<void> | void) {
    await callback(this.context);
  },
};

describe("Word", function () {
  it("Inserts text", async function () {
    const wordMock: OfficeMockObject = new OfficeMockObject(WordMockData); // Mocking the host specific namespace
    global.Word = wordMock as any;

    await insertText("Hello Word");

    wordMock.context.document.body.paragraph.load("text");
    await wordMock.context.sync();

    assert.strictEqual(wordMock.context.document.body.paragraph.text, "Hello Word");
  });
});
