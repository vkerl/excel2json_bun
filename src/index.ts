#!/usr/bin/env bun
/**
 * excel2json 命令行入口
 *
 * 用法：
 *   bun run src/index.ts --input data.xlsx --output ./output
 *   bun run src/index.ts data.xlsx -o ./output
 *   bun run src/index.ts data.xlsx -o ./output --sheet Player,Items   # 只导出指定 Sheet
 *   bun run src/index.ts data.xlsx -o ./output --compact              # 紧凑 JSON（单行）
 */
import * as fs from "fs";
import * as path from "path";
import { parseArgs } from "./args";
import { readWorkbook, readAllSheetRows } from "./excel";
import { convertWorkbook, makeResolver } from "./converter";

function printUsage() {
  console.log(`
excel2json - 把 Excel 转为 JSON

用法:
  excel2json --input <文件.xlsx> [--output <目录>] [--sheet 名字1,名字2] [--compact]

参数:
  -i, --input    输入的 Excel 文件路径（必填，也可作为第一个位置参数）
  -o, --output   输出文件或目录：
                 -o ./a/b.json   直接写单个文件（多 Sheet 会合并为按名 key 的对象）
                 -o ./out        目录模式，每个 Sheet 一个 .json（默认 ./output）
  -s, --sheet    只导出指定 Sheet，逗号分隔；可加 :行范围 限制行（1-based）
                例: --sheet Player,Bag:1-2
  -n, --name     自定义输出文件名，格式 Sheet=文件.json（逗号分隔，可省 .json）
                例: --name Player=player_cfg.json,Bag=bag.json
      --compact  输出紧凑（不缩进）的 JSON
  -h, --help     显示本帮助

Excel 格式约定（每个 Sheet）:
  第 1 行  注释行（不导出）
  第 2 行  数据类型: number | string | bool | []number | []string | []bool | []object | comment
                （comment = 注释列，仅人工标注，不导出）
  第 3 行  字段名（导出 JSON 的 key；comment 列的字段名可留空）
  第 4 行+ 数据
  []object 单元格支持 #include 跨表引用，并可按行/条件过滤:
    #include Items                      整个 Sheet
    #include Bag by pid=id              关联：Bag.pid == 本行 id
    #include Bag by type="weapon"       字面量：Bag.type == "weapon"
    #include Items where price > 100    where 条件（= != > >= < <= in）
    #include Items rows 1,3             只取第 1、3 条记录
    #include Bag by pid=id rows 1-3     子句可组合
`);
}

function sanitize(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, "_").trim() || "sheet";
}

function main() {
  const args = parseArgs(Bun.argv.slice(2));

  if (args.help) {
    printUsage();
    process.exit(0);
  }

  const input = args.input || (args as any).i || args._[0];
  if (!input) {
    console.error("❌ 缺少输入文件，使用 --help 查看用法");
    process.exit(1);
  }

  const output = args.output || (args as any).o || "./output";
  // 支持 --sheet Player,Bag:1-2 形式（冒号后为行范围，1-based）
  const sheetSpecs = args.sheet || (args as any).s
    ? String(args.sheet || (args as any).s)
        .split(",")
        .map((s) => {
          const [name, rowSpec] = s.split(":");
          return {
            name: name.trim(),
            rowSpec: rowSpec ? rowSpec.trim() : undefined,
          };
        })
        .filter((s) => s.name)
    : undefined;
  const compact = !!args.compact;

  // 输出文件名映射：SheetName=file.json（可省略 .json）
  const nameArg = args.name || (args as any).n;
  const nameMap = nameArg ? parseNameMap(String(nameArg)) : undefined;

  if (!fs.existsSync(input)) {
    console.error(`❌ 输入文件不存在: ${input}`);
    process.exit(1);
  }

  console.log(`读取: ${input}`);
  const wb = readWorkbook(input);
  const sheetRows = readAllSheetRows(wb);
  const resolve = makeResolver((name) => {
    const rows = sheetRows.get(name);
    if (!rows) throw new Error(`找不到被 #include 引用的 Sheet "${name}"`);
    return rows;
  });

  const result = convertWorkbook(sheetRows, resolve, { sheets: sheetSpecs });

  const written: string[] = [];
  if (output.toLowerCase().endsWith(".json")) {
    // 单文件模式：-o ./a/b.json
    fs.mkdirSync(path.dirname(path.resolve(output)), { recursive: true });
    const names = Object.keys(result);
    const payload = names.length === 1 ? Object.values(result)[0] : result;
    fs.writeFileSync(output, JSON.stringify(payload, null, compact ? 0 : 2));
    written.push(output);
    if (names.length > 1) {
      console.warn("  ⚠ 检测到多个 Sheet，已合并写入单个文件（以 Sheet 名为 key）");
    }
  } else {
    // 目录模式：每个 Sheet 一个文件
    fs.mkdirSync(output, { recursive: true });
    const usedFiles = new Set<string>();
    for (const [name, data] of Object.entries(result)) {
      const target =
        nameMap && nameMap.has(name) ? nameMap.get(name)! : `${sanitize(name)}.json`;
      if (usedFiles.has(target)) {
        console.warn(`  ⚠ 输出文件名冲突，已覆盖: ${target}`);
      }
      usedFiles.add(target);
      const file = path.join(output, target);
      fs.writeFileSync(file, JSON.stringify(data, null, compact ? 0 : 2));
      written.push(file);
    }
  }

  console.log(`✅ 已导出 ${written.length} 个文件到 ${output}`);
  for (const f of written) console.log(`   - ${f}`);
}

/** 解析 --name 映射：SheetName=file.json，逗号分隔，可省略 .json */
function parseNameMap(s: string): Map<string, string> {
  const m = new Map<string, string>();
  for (const part of s.split(",")) {
    const eq = part.indexOf("=");
    if (eq < 0) continue;
    const sheet = part.slice(0, eq).trim();
    let file = part.slice(eq + 1).trim();
    if (!file) continue;
    if (!file.toLowerCase().endsWith(".json")) file += ".json";
    m.set(sheet, file);
  }
  return m;
}

main();
