Office.onReady(() => {
});

async function addRubi(event: Office.AddinCommands.Event) {
  Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();
    const text = range.text;
    const rubitext = "かくにん";
    const field = range.insertField(
      Word.InsertLocation.replace,
      Word.FieldType.eq,
      "\\* jc2 \\* hps10 \\o(\\s\\up9(" + rubitext + ")," + text + ")",
      true
    );
    await context.sync();
  });
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
      ? window
      : typeof global !== "undefined"
        ? global
        : undefined;
}

const g = getGlobal();

Office.actions.associate("addRubi", addRubi);