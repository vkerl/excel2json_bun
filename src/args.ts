export interface Args {
  input?: string;
  output?: string;
  sheet?: string;
  compact?: boolean;
  help?: boolean;
  _: string[];
}

/**
 * 极简命令行参数解析（无第三方依赖）。
 *   --key value   或   --key=value
 *   -k value      或   -k=value
 *   --flag / -f   作为布尔开关
 *   其余作为位置参数（第一个位置参数视为 input）
 */
export function parseArgs(argv: string[]): Args {
  const args: Args = { _: [] };

  const set = (key: string, value: string | boolean) => {
    (args as any)[key] = value;
  };

  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "-h" || a === "--help") {
      args.help = true;
      continue;
    }

    let key: string | null = null;
    let inline: string | undefined;

    if (a.startsWith("--")) {
      const body = a.slice(2);
      const eq = body.indexOf("=");
      if (eq >= 0) {
        key = body.slice(0, eq);
        inline = body.slice(eq + 1);
      } else {
        key = body;
      }
    } else if (a.startsWith("-") && a.length > 1) {
      const body = a.slice(1);
      const eq = body.indexOf("=");
      if (eq >= 0) {
        key = body.slice(0, eq);
        inline = body.slice(eq + 1);
      } else {
        key = body;
      }
    } else {
      args._.push(a);
      continue;
    }

    if (inline !== undefined) {
      set(key!, inline);
      continue;
    }
    const next = argv[i + 1];
    if (next !== undefined && !next.startsWith("-")) {
      set(key!, next);
      i++;
    } else {
      set(key!, true);
    }
  }

  return args;
}
