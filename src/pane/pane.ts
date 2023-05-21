function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
      ? window
      : global;
}

const g = getGlobal();

Office.onReady(() => {
  g.document.getElementById("button")!.addEventListener("click", () => tryCatch(addRubi), false);
  g.document.getElementById("nazo")!.addEventListener("click", () => tryCatch(nazonoSpace), false);
});

async function addRubi() {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();
    const rubidata = rubi(range.text);
    let nowRange = range;
    for (const iterator of rubidata) {
      const text = iterator.s;
      const rubitext = iterator.r;
      const code = "\\* jc2 \\* hps10 \\o(\\s\\up9(" + rubitext + ")," + text + ")";
      console.log(code);
      nowRange.insertField(
        Word.InsertLocation.before,
        Word.FieldType.eq,
        code.trim(),
        true
      );
    }
    range.clear();
    await context.sync();
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}


async function nazonoSpace() {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("fields");
    await context.sync();
    for (const iterator of range.fields.items) {
      iterator.code = iterator.code.trim()
    }
    await context.sync();
  });
}
