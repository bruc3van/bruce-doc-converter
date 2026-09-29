/**
 * 结构化转换诊断，与 Python 侧约定一致：
 * severity 为 'warning'（内容缺失或降级，严格模式拒绝）或 'info'（提示，不阻止导出）。
 */
const MAX_DIAGNOSTICS = 100;

class Diagnostics {
  constructor(limit = MAX_DIAGNOSTICS) {
    this.limit = limit;
    this.items = [];
    this.omitted = 0;
  }

  add(code, message, severity = 'warning', line) {
    const item = { code, severity, message: String(message).slice(0, 300), ...(Number.isInteger(line) ? { line } : {}) };
    if (this.items.length < this.limit) {
      this.items.push(item);
      return;
    }
    // Keep the list bounded while never hiding that a blocking diagnostic occurred.
    this.omitted += 1;
    const previous = this.items[this.limit - 1];
    this.items[this.limit - 1] = {
      code: 'DIAGNOSTICS_TRUNCATED',
      severity: previous.severity === 'warning' || severity === 'warning' ? 'warning' : 'info',
      message: `另有 ${this.omitted + 1} 条诊断被省略。`
    };
  }

  /** Human-readable messages of blocking diagnostics, for the legacy warnings field. */
  warnings() {
    return this.items
      .filter(item => item.severity === 'warning')
      .map(item => (item.line ? `${item.message}（第 ${item.line} 行）` : item.message));
  }
}

module.exports = { Diagnostics };
