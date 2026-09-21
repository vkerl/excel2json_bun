/**
 * 生成测试用 sample.xlsx，演示 3 行表头 + #include 跨表引用 / 关联 / 行过滤。
 * 运行: bun run scripts/sample.ts
 */
import * as XLSX from "xlsx";

// Sheet: Items（被 Player 通过 #include 整表引用）
const items: any[][] = [
  ["物品ID", "物品名", "价格", "是否稀有"],
  ["number", "string", "number", "bool"],
  ["id", "name", "price", "rare"],
  [1, "铁剑", 100, false],
  [2, "木盾", 80, false],
  [3, "屠龙刀", 9999, true],
];

// Sheet: Bag（每个玩家自己的背包，pid 关联 Player.id）
const bag: any[][] = [
  ["玩家ID", "物品名", "数量"],
  ["number", "string", "number"],
  ["pid", "item", "count"],
  [1001, "铁剑", 2],
  [1001, "木盾", 1],
  [1002, "屠龙刀", 1],
  [1003, "长枪", 5],
];

// Sheet: Player（含 comment 注释列，不导出）
const player: any[][] = [
  ["玩家ID", "名字", "等级列表", "是否VIP", "全部物品", "我的背包", "前两件物品", "标签", "备注"],
  [
    "number",
    "string",
    "[]number",
    "bool",
    "[]object",
    "[]object",
    "[]object",
    "[]string",
    "comment",
  ],
  ["id", "name", "levels", "vip", "items", "bag", "topItems", "tags", ""],
  [1001, "张三", "1,2,3", true, "#include Items", "#include Bag by pid=id", "#include Items rows 1,2", "战士,法师", "老玩家"],
  [1002, "李四", "5,10", false, "#include Items", "#include Bag by pid=id", "#include Items rows 1,2", "坦克", "新手"],
  [1003, "王五", "7", true, "#include Items", "#include Bag by pid=id", "#include Items rows 1,2", "刺客,弓手", "回归玩家"],
];

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(items), "Items");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(bag), "Bag");
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(player), "Player");
XLSX.writeFile(wb, "sample.xlsx");
console.log("已生成 sample.xlsx");
