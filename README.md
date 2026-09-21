# excel2json

使用 [Bun](https://bun.sh) 编写的 Excel → JSON 转换工具，支持构建 macOS / Windows（及 Linux）独立可执行文件。

## 功能

1. **第 1 行：注释行** —— 仅用于人工标注，不导出。
2. **第 2 行：数据类型行** —— 标注每列类型，支持：
   - `number` / `string` / `bool`
   - `[]number` / `[]string` / `[]bool` / `[]object`
   - `comment` —— 注释列，仅用于人工标注，**不导出**到 JSON（字段名可留空）
3. **第 3 行：字段名** —— 导出 JSON 的 key，数组单元格用**英文逗号**分隔元素。
4. **`[]object` 跨表引用** —— 单元格内写 `#include` 语法，把另一个 Sheet 的数据作为对象数组内联进来，并支持按行/条件过滤：

   | 写法 | 含义 |
   | ---- | ---- |
   | `#include Items` | 导入整个 Sheet |
   | `#include("Items")` / `#include 'Items'` | 同上（带引号/括号） |
   | `#include Bag by pid=id` | **关联**：取 `Bag.pid` 等于「本行 `id` 字段」的那些行 |
   | `#include Bag by type="weapon"` | 字面量过滤：`Bag.type == "weapon"`（加引号强制当字面量） |
   | `#include Items where price > 100` | `where` 条件（`= != > >= < <= in`） |
   | `#include Items rows 1,3` | 只取第 1、3 条记录 |
   | `#include Items rows 2-5` | 只取第 2~5 条记录 |
   | `#include Bag by pid=id rows 1-3` | 子句可任意组合 |

   > 关联说明：`by field=主表字段` 中，若右侧名字恰好是本行已存在的字段名，则取本行该字段的值进行匹配；若用引号包裹（`"..."`）则强制当作字面量。匹配比较对 number/string/bool 都做了兼容。
   > 退化支持：单元格直接写 JSON 数组/对象。

## 安装与运行（开发模式）

```bash
bun install
bun run scripts/sample.ts        # 生成 sample.xlsx
bun run src/index.ts --input sample.xlsx --output ./output
```

## 构建各平台可执行文件

```bash
bun install
bun run build          # 一次性构建 mac / win / linux 全部目标
# 或单独构建：
bun run build:mac      # macOS arm64 + x64
bun run build:win      # Windows x64 + arm64
bun run build:linux    # Linux x64
```

产物在 `dist/` 目录下，例如 `excel2json-windows-x64.exe`、`excel2json-darwin-arm64`。

> 说明：Bun 的 `bun build --compile` 支持跨平台编译，因此在 macOS 上也能直接产出 Windows 可执行文件。若某些环境无法跨编译，可在对应平台上执行对应脚本。

## 命令行用法

```
excel2json --input data.xlsx [--output ./output] [--sheet 名字1,名字2] [--compact]

  -i, --input   输入 Excel 路径（必填，也可作第一个位置参数）
  -o, --output  输出目录（默认 ./output）
  -s, --sheet   只导出指定 Sheet，逗号分隔
      --compact 紧凑（单行）JSON
  -h, --help    帮助
```

## Excel 示例

| 注释 | 玩家ID | 名字 | 等级列表 | 是否VIP | 背包物品 | 标签 |
| ---- | ------ | ---- | -------- | ------- | -------- | ---- |
| 类型 | number | string | []number | bool | []object | []string |
| 字段 | id | name | levels | vip | items | tags |
| 数据 | 1001 | 张三 | 1,2,3 | TRUE | #include Items | 战士,法师 |

导出的 `Player.json` 示例（`items` 为整表引用，`bag` 为按本行 `id` 关联的背包，`topItems` 为行过滤）：

```json
[
  {
    "id": 1001,
    "name": "张三",
    "levels": [1, 2, 3],
    "vip": true,
    "items": [
      { "id": 1, "name": "铁剑", "price": 100, "rare": false },
      { "id": 2, "name": "木盾", "price": 80, "rare": false },
      { "id": 3, "name": "屠龙刀", "price": 9999, "rare": true }
    ],
    "bag": [
      { "pid": 1001, "item": "铁剑", "count": 2 },
      { "pid": 1001, "item": "木盾", "count": 1 }
    ],
    "topItems": [
      { "id": 1, "name": "铁剑", "price": 100, "rare": false },
      { "id": 2, "name": "木盾", "price": 80, "rare": false }
    ],
    "tags": ["战士", "法师"]
  }
]
```

## 只导出某个 Sheet 的某几行

通过 `-s/--sheet` 加上 `:行范围`（1-based）限制顶层导出的行：

```bash
excel2json -i sample.xlsx -o ./output -s "Player:1-2"
# 只导出 Player 的第 1、2 条记录
```

## 自定义输出文件名

最简单的方式：直接用 `-o` 指定输出文件路径（支持子目录，自动创建）：

```bash
excel2json -i sample.xlsx -o ./a/b.json
# 若只有一个 Sheet，b.json 即该 Sheet 的数据（数组）
# 若有多个 Sheet，b.json 为按 Sheet 名分 key 合并的对象
```

如果需要按 Sheet 分别命名，再用 `-o` 指向目录 + `-n/--name` 映射：

```bash
excel2json -i sample.xlsx -o ./output -n "Player=player_cfg.json,Bag=bag.json"
# Player -> player_cfg.json，Bag -> bag.json，其余仍用 Sheet 名
```

> 注意：`-o` 以 `.json` 结尾视为单文件；否则视为目录。多个 Sheet 映射到同一文件名会互相覆盖（打印警告）。

## 转换规则

- 空单元格：标量 → `null`/`false`/`""`，数组 → `[]`。
- 完全为空的行会被跳过。
- 数组单元格以英文逗号分隔，自动去除首尾空格。
- `bool` 接受：`true/1/yes/y/是` 为 true，其余为 false。
- `#include` 支持缓存与循环引用检测（避免互相引用造成死循环）。
