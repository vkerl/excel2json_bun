/**
 * Excel -> JSON 核心转换逻辑（纯函数，不依赖任何 Excel 解析库）。
 *
 * Excel 表格约定（每个 Sheet）：
 *   第 1 行：注释行（用户标注，不导出）
 *   第 2 行：数据类型行，支持：
 *            number | string | bool
 *            []number | []string | []bool | []object
 *   第 3 行：字段名（导出后的 JSON 字段 key）
 *   第 4 行起：数据行
 *
 * []object 类型支持 `#include` 语法，把另一个 Sheet 的数据作为对象数组内联进来：
 *   #include SheetName                         导出整个 Sheet
 *   #include Sheet by field=value              按字面量过滤（field == value）
 *   #include Sheet by field=mainField          关联过滤：field 等于「本行」字段 mainField 的值
 *   #include Sheet where field > 5             where 条件（= != > >= < <= in）
 *   #include Sheet rows 1,3                    只取第 1、3 条记录
 *   #include Sheet rows 2-5                    只取第 2~5 条记录
 *   以上子句可组合，例如：#include Bag by pid=id rows 1-3
 *
 * 说明：mainField / value 若用引号包裹（"..." 或 '...'）则强制当作字面量；
 * 否则若该名字恰好是本行已存在的字段名，则取本行该字段的值（关联）。
 */

export type CellType =
  | "number"
  | "string"
  | "bool"
  | "[]number"
  | "[]string"
  | "[]bool"
  | "[]object"
  | "comment";

export type ResolveFn = (sheetName: string) => any[];

const ALLOWED: CellType[] = [
  "number",
  "string",
  "bool",
  "[]number",
  "[]string",
  "[]bool",
  "[]object",
  "comment", // 注释列：仅用于人工标注，不导出
];

interface IncludeSpec {
  sheet: string;
  joinField?: string;
  joinKey?: string; // 关联本行字段名（未加引号时）
  joinLiteral?: any; // 显式字面量（加引号时）
  whereField?: string;
  whereOp?: string;
  whereKey?: string;
  whereLiteral?: any;
  rowSpec?: string;
}

function normalizeType(raw: unknown): CellType {
  const t = String(raw ?? "").trim().toLowerCase();
  if ((ALLOWED as string[]).includes(t)) return t as CellType;
  throw new Error(
    `不支持的数据类型: "${raw}"（支持: ${ALLOWED.join(", ")}）`,
  );
}

function toNumber(raw: any): number | null {
  if (typeof raw === "number") return raw;
  if (typeof raw === "boolean") return raw ? 1 : 0;
  const s = String(raw).replace(/,/g, "").trim();
  if (s === "") return null;
  const n = Number(s);
  return Number.isNaN(n) ? null : n;
}

function toBool(raw: any): boolean {
  if (typeof raw === "boolean") return raw;
  const s = String(raw).trim().toLowerCase();
  return s === "true" || s === "1" || s === "yes" || s === "y" || s === "是";
}

function splitArray(raw: any): string[] {
  return String(raw)
    .split(",")
    .map((s) => s.trim())
    .filter((s) => s !== "");
}

function isEmptyValue(v: any): boolean {
  if (v === null || v === undefined) return true;
  if (typeof v === "string" && v === "") return true;
  if (Array.isArray(v) && v.length === 0) return true;
  return false;
}

function isQuoted(s: string): boolean {
  return (
    (s.startsWith('"') && s.endsWith('"')) ||
    (s.startsWith("'") && s.endsWith("'"))
  );
}

function stripQuotes(s: string): string {
  if (isQuoted(s)) return s.slice(1, -1);
  return s;
}

function coerceLiteral(s: string): any {
  const t = stripQuotes(s).trim();
  if (t === "true") return true;
  if (t === "false") return false;
  if (t !== "" && !isNaN(Number(t))) return Number(t);
  return t;
}

function splitFirst(s: string, ch: string): [string, string] {
  const i = s.indexOf(ch);
  if (i < 0) return [s, ""];
  return [s.slice(0, i), s.slice(i + 1)];
}

function tokenize(s: string): string[] {
  const tokens: string[] = [];
  let cur = "";
  let quote: string | null = null;
  for (const ch of s) {
    if (quote) {
      if (ch === quote) quote = null;
      else cur += ch;
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (/\s/.test(ch)) {
      if (cur) {
        tokens.push(cur);
        cur = "";
      }
    } else {
      cur += ch;
    }
  }
  if (cur) tokens.push(cur);
  return tokens;
}

function parseInclude(cell: string): IncludeSpec {
  const body = cell.trim();
  const m = body.match(/^#include\s*([\s\S]*)$/);
  const rest = m ? m[1] : body;
  const spec: IncludeSpec = { sheet: "" };
  const tokens = tokenize(rest);
  if (tokens.length === 0) return spec;

  let sheet = tokens.shift()!;
  sheet = sheet.replace(/^\(/, "").replace(/\)$/, "");
  spec.sheet = stripQuotes(sheet);

  while (tokens.length) {
    const t = tokens.shift()!;
    if (t === "by") {
      const kv = tokens.shift() || "";
      const [f, v] = splitFirst(kv, "=");
      spec.joinField = f.trim();
      if (isQuoted(v)) spec.joinLiteral = coerceLiteral(v);
      else spec.joinKey = v.trim();
    } else if (t === "where") {
      const f = (tokens.shift() || "").trim();
      const op = (tokens.shift() || "=").trim();
      const v = tokens.shift() || "";
      spec.whereField = f;
      spec.whereOp = op;
      if (isQuoted(v)) spec.whereLiteral = coerceLiteral(v);
      else spec.whereKey = v.trim();
    } else if (t === "rows") {
      spec.rowSpec = (tokens.shift() || "").trim();
    }
    // 未知 token 忽略
  }
  return spec;
}

function valuesEqual(a: any, b: any): boolean {
  if (a === b) return true;
  if (a === null || a === undefined || b === null || b === undefined)
    return false;
  if (typeof a === "number" && typeof b === "number") return a === b;
  if (typeof a === "boolean" || typeof b === "boolean")
    return Boolean(a) === Boolean(b);
  return String(a).trim() === String(b).trim();
}

function num(v: any): number {
  const n = typeof v === "number" ? v : Number(String(v).replace(/,/g, ""));
  return Number.isNaN(n) ? 0 : n;
}

function compareOp(actual: any, op: string, expected: any): boolean {
  switch (op) {
    case "=":
    case "==":
      return valuesEqual(actual, expected);
    case "!=":
    case "<>":
      return !valuesEqual(actual, expected);
    case ">":
      return num(actual) > num(expected);
    case ">=":
      return num(actual) >= num(expected);
    case "<":
      return num(actual) < num(expected);
    case "<=":
      return num(actual) <= num(expected);
    case "in": {
      const list = Array.isArray(expected)
        ? expected
        : String(expected)
            .split(",")
            .map((s) => coerceLiteral(s.trim()));
      return list.some((v) => valuesEqual(actual, v));
    }
    default:
      return valuesEqual(actual, expected);
  }
}

function parseRowSpec(spec: string): (idx: number) => boolean {
  const set = new Set<number>();
  const ranges: [number, number][] = [];
  for (const part of spec.split(",")) {
    const p = part.trim();
    if (!p) continue;
    const dash = p.indexOf("-");
    if (dash > 0) {
      const a = parseInt(p.slice(0, dash), 10);
      const b = parseInt(p.slice(dash + 1), 10);
      if (!isNaN(a) && !isNaN(b)) ranges.push([Math.min(a, b), Math.max(a, b)]);
    } else {
      const n = parseInt(p, 10);
      if (!isNaN(n)) set.add(n);
    }
  }
  return (idx: number) =>
    set.has(idx) || ranges.some(([a, b]) => idx >= a && idx <= b);
}

function resolveRef(
  key: string | undefined,
  literal: any,
  hasLiteral: boolean,
  mainRow: Record<string, any>,
): any {
  if (hasLiteral) return literal;
  if (key === undefined) return undefined;
  if (Object.prototype.hasOwnProperty.call(mainRow, key)) return mainRow[key];
  return coerceLiteral(key);
}

function applyInclude(
  spec: IncludeSpec,
  resolve: ResolveFn,
  mainRow: Record<string, any>,
): any[] {
  let arr = resolve(spec.sheet);

  if (spec.joinField) {
    const val = resolveRef(
      spec.joinKey,
      spec.joinLiteral,
      spec.joinLiteral !== undefined,
      mainRow,
    );
    arr = arr.filter((o) => valuesEqual(o[spec.joinField], val));
  }

  if (spec.whereField) {
    const val = resolveRef(
      spec.whereKey,
      spec.whereLiteral,
      spec.whereLiteral !== undefined,
      mainRow,
    );
    arr = arr.filter((o) => compareOp(o[spec.whereField], spec.whereOp || "=", val));
  }

  if (spec.rowSpec) {
    const pred = parseRowSpec(spec.rowSpec);
    arr = arr.filter((_, idx) => pred(idx + 1));
  }

  return arr;
}

function convertObjectCell(
  raw: any,
  resolve: ResolveFn,
  mainRow: Record<string, any>,
): any[] {
  const str = String(raw).trim();
  if (str.toLowerCase().startsWith("#include")) {
    const spec = parseInclude(str);
    if (!spec.sheet) return [];
    return applyInclude(spec, resolve, mainRow);
  }
  // 退化支持：单元格内直接写 JSON 数组 / 对象
  try {
    const parsed = JSON.parse(str);
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch {
    return [];
  }
}

function convertCell(
  raw: any,
  type: CellType,
  resolve: ResolveFn,
  mainRow: Record<string, any>,
): any {
  if (raw === undefined || raw === null || raw === "") {
    if (type.startsWith("[]")) return [];
    if (type === "bool") return false;
    if (type === "number") return null;
    return "";
  }
  switch (type) {
    case "number":
      return toNumber(raw);
    case "string":
      return String(raw);
    case "bool":
      return toBool(raw);
    case "[]number":
      return splitArray(raw).map(toNumber);
    case "[]string":
      return splitArray(raw).map((s) => s);
    case "[]bool":
      return splitArray(raw).map(toBool);
    case "[]object":
      return convertObjectCell(raw, resolve, mainRow);
    default:
      return raw;
  }
}

export interface ConvertSheetOptions {
  /** 顶层导出时只保留第 N 条记录（1-based，按导出顺序计）。include 内部不使用。 */
  dataRowFilter?: (recordIndex: number) => boolean;
}

/**
 * 把一个 Sheet 的数据行转换为对象数组。
 * @param rows    该 Sheet 的全部行（array of array）
 * @param resolve 通过名字解析另一个 Sheet 的对象数组（用于 #include）
 * @param options 可选的行过滤（用于 --sheet Name:rows）
 */
export function convertSheet(
  rows: any[][],
  resolve: ResolveFn,
  options: ConvertSheetOptions = {},
): any[] {
  if (!rows || rows.length < 3) {
    console.warn("  ⚠ 数据不足 3 行（需要注释行/类型行/字段行），已跳过该 Sheet");
    return [];
  }

  const typeRow = rows[1] ?? [];
  const fieldRow = rows[2] ?? [];

  const maxCol = Math.max(fieldRow.length, typeRow.length);
  const fields: string[] = [];
  const types: CellType[] = [];

  for (let j = 0; j < maxCol; j++) {
    const field = String(fieldRow[j] ?? "").trim();
    const type = normalizeType(typeRow[j]);
    fields[j] = field;
    types[j] = type;
  }

  const result: any[] = [];
  let recordIndex = 0;
  for (let i = 3; i < rows.length; i++) {
    const dataRow = rows[i] ?? [];

    // 先收集本行各字段原始值，供 #include 关联引用
    const rowRaw: Record<string, any> = {};
    for (let j = 0; j < fields.length; j++) {
      if (fields[j]) rowRaw[fields[j]] = dataRow[j];
    }

    const obj: Record<string, any> = {};
    let hasContent = false;
    for (let j = 0; j < fields.length; j++) {
      const field = fields[j];
      if (types[j] === "comment") continue; // 注释列：不导出
      if (!field) continue; // 没有字段名的列直接忽略
      const val = convertCell(dataRow[j], types[j], resolve, rowRaw);
      obj[field] = val;
      if (!isEmptyValue(val)) hasContent = true;
    }
    if (!hasContent) continue; // 跳过完全为空的行

    recordIndex++;
    if (options.dataRowFilter && !options.dataRowFilter(recordIndex)) continue;
    result.push(obj);
  }
  return result;
}

/**
 * 构造一个 resolve 函数：根据 Sheet 名返回其对象数组，
 * 内置缓存与循环引用检测（防止 #include 互相引用导致死循环）。
 * @param provider 由外部提供：给定 Sheet 名返回其全部行
 */
export function makeResolver(provider: (name: string) => any[][]): ResolveFn {
  const cache = new Map<string, any[]>();
  const visiting = new Set<string>();

  function resolve(name: string): any[] {
    if (cache.has(name)) return cache.get(name)!;
    if (visiting.has(name)) {
      throw new Error(`检测到 #include 循环引用: "${name}"`);
    }
    visiting.add(name);
    const rows = provider(name);
    const arr = convertSheet(rows, resolve);
    visiting.delete(name);
    cache.set(name, arr);
    return arr;
  }

  return resolve;
}

export interface SheetExportSpec {
  name: string;
  /** 只导出该 Sheet 的第 N 条记录（1-based，逗号/短横，例如 "1,3" 或 "2-5"） */
  rowSpec?: string;
}

export interface ConvertOptions {
  /** 只导出指定 Sheet；不传则导出全部 */
  sheets?: SheetExportSpec[];
}

/**
 * 转换整个工作簿。
 * @param sheetRows  Map<Sheet名, 全部行>
 * @param resolve    #include 解析函数
 * @param options    导出选项（可限制 Sheet 与行）
 */
export function convertWorkbook(
  sheetRows: Map<string, any[][]>,
  resolve: ResolveFn,
  options: ConvertOptions = {},
): Record<string, any[]> {
  const specs =
    options.sheets && options.sheets.length
      ? options.sheets
      : Array.from(sheetRows.keys()).map((name) => ({ name }));

  const out: Record<string, any[]> = {};
  for (const spec of specs) {
    const rows = sheetRows.get(spec.name);
    if (!rows) {
      console.warn(`  ⚠ 找不到 Sheet "${spec.name}"，已跳过`);
      continue;
    }
    const filter = spec.rowSpec ? parseRowSpec(spec.rowSpec) : undefined;
    out[spec.name] = convertSheet(rows, resolve, { dataRowFilter: filter });
  }
  return out;
}
