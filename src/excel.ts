import * as fs from "fs";
import * as XLSX from "xlsx";

/**
 * 读取 Excel 文件为工作簿对象。
 * 注意：在 Bun 编译出的独立二进制中，xlsx 自带的 XLSX.readFile 无法访问
 * 宿主机文件系统，因此这里自己用 fs.readFileSync 读取字节，再交给 XLSX.read。
 */
export function readWorkbook(file: string): XLSX.WorkBook {
  const buf = fs.readFileSync(file);
  return XLSX.read(buf, { type: "buffer", cellDates: true });
}

/**
 * 把一个工作簿拆成 Map<Sheet名, 全部行(array of array)>。
 * raw:true 保留原始类型（数字/布尔/字符串），空单元格用 "" 占位。
 */
export function readAllSheetRows(wb: XLSX.WorkBook): Map<string, any[][]> {
  const map = new Map<string, any[][]>();
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      defval: "",
      raw: true,
    }) as any[][];
    map.set(name, rows);
  }
  return map;
}
