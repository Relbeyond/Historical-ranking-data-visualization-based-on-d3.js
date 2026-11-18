(() => {
  const input = document.querySelector("#inputfile");
  const sheet = document.querySelector("#worksheet");
  const sheetLabel = document.querySelector("#worksheet-label");
  const status = document.querySelector("#status");
  const fields = ["name", "type", "value", "date"];
  let workbook;
  let filename;
  let stopAnimation;
  let selection = 0;

  function report(message, error = false) {
    status.textContent = message;
    status.dataset.error = error;
  }

  function display(rows, source) {
    const headers = (rows[0] || []).map(value => String(value).trim().toLowerCase());
    const columns = fields.map(field => headers.indexOf(field));
    if (columns.includes(-1)) {
      throw new Error("表头需包含 name、type、value、date 四列。");
    }

    const data = [];
    for (const [index, row] of rows.slice(1).entries()) {
      if (row.every(value => String(value ?? "").trim() === "")) continue;
      const [name, type, value, date] = columns.map(column => {
        const cell = row[column];
        return cell instanceof Date ? cell.toISOString().slice(0, 10) : String(cell ?? "").trim();
      });
      if (!name || !date || !value || !Number.isFinite(Number(value))) {
        throw new Error(`第 ${index + 2} 行需要填写名称、日期和有效数值。`);
      }
      data.push({name, type, value: Number(value), date});
    }
    if (!data.length) throw new Error("表格没有可绘制的数据。");

    if (stopAnimation) stopAnimation();
    d3.select("svg").selectAll("*").remove();
    stopAnimation = draw(data);
    report(`${source} · ${data.length} 行`);
  }

  function displaySheet() {
    try {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheet.value], {
        header: 1,
        raw: true,
        defval: ""
      });
      display(rows, `${filename} / ${sheet.value}`);
    } catch (error) {
      report(error.message, true);
    }
  }

  sheet.addEventListener("change", displaySheet);
  input.addEventListener("change", async () => {
    const file = input.files[0];
    if (!file) return;
    const currentSelection = ++selection;
    input.value = "";
    sheet.hidden = sheetLabel.hidden = true;
    sheet.replaceChildren();
    workbook = undefined;
    report(`正在读取 ${file.name}…`);

    try {
      const buffer = await file.arrayBuffer();
      if (currentSelection !== selection) return;
      if (/\.csv$/i.test(file.name)) {
        let content;
        try {
          content = new TextDecoder(config.encoding, {fatal: true}).decode(buffer);
        } catch {
          content = new TextDecoder("gb18030").decode(buffer);
        }
        display(d3.csvParseRows(content), file.name);
      } else if (/\.xlsx?$/i.test(file.name)) {
        workbook = XLSX.read(buffer, {cellDates: true});
        filename = file.name;
        for (const name of workbook.SheetNames) sheet.add(new Option(name, name));
        sheet.hidden = sheetLabel.hidden = workbook.SheetNames.length < 2;
        displaySheet();
      } else {
        throw new Error("请选择 CSV、XLSX 或 XLS 文件。");
      }
    } catch (error) {
      if (currentSelection === selection) report(`导入失败：${error.message}`, true);
    }
  });
})();
