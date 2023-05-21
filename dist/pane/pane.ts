function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
      ? window
      : global;
}

const g = getGlobal();

Office.onReady(() => {
  const button: HTMLElement = <HTMLElement>g.document.getElementById("button");
  button.addEventListener("click", () => tryCatch(addRubi), false);
});

async function addRubi() {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();
    const rubidata = rubi(range.text);
    range.clear();
    await context.sync();
    for (const iterator of rubidata) {
      const text = iterator.s;
      const rubitext = iterator.r;
      const field = range.insertField(
        Word.InsertLocation.replace,
        Word.FieldType.eq,
        "\\* jc2 \\* hps10 \\o(\\s\\up9(" + rubitext + ")," + text + ")",
        true
      );
      await context.sync();
    }
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
