// Excel Formula Cards Pro — Formula Database
// Each formula: { id, zh, desc, cat, diff, ver, syntax, args, ex, itype, ivals, ifn, tips, rel }
// diff: 'b' = beginner, 'i' = intermediate, 'a' = advanced
// ver: 'all', '2016', '2019', '365'
// itype: 'nums' | 'text1' | 'text2' | 'text3' | 'date' | 'table' | 'none'

window.FORMULA_COMPUTE = {
  sum:        v => v.reduce((a, b) => a + b, 0),
  avg:        v => +(v.reduce((a, b) => a + b, 0) / v.length).toFixed(4),
  max:        v => Math.max(...v),
  min:        v => Math.min(...v),
  count:      v => v.length,
  product:    v => v.reduce((a, b) => a * b, 1),
  round:      ([n, d = 2]) => +parseFloat(n).toFixed(+d),
  roundup:    ([n, d = 0]) => { const f = Math.pow(10, +d); return Math.ceil(+n * f) / f; },
  rounddown:  ([n, d = 0]) => { const f = Math.pow(10, +d); return Math.floor(+n * f) / f; },
  int_:       ([n]) => Math.floor(+n),
  mod:        ([n, d]) => +(+n % +d).toFixed(6),
  abs:        ([n]) => Math.abs(+n),
  power:      ([n, e]) => +Math.pow(+n, +e).toFixed(4),
  sqrt:       ([n]) => +Math.sqrt(+n).toFixed(4),
  ceiling:    ([n, s = 1]) => Math.ceil(+n / +s) * +s,
  floor_:     ([n, s = 1]) => Math.floor(+n / +s) * +s,
  large:      (v, k = 2) => [...v].sort((a, b) => b - a)[+k - 1],
  small:      (v, k = 2) => [...v].sort((a, b) => a - b)[+k - 1],
  // text
  len:        ([t]) => t.length,
  upper:      ([t]) => t.toUpperCase(),
  lower:      ([t]) => t.toLowerCase(),
  proper:     ([t]) => t.replace(/\w\S*/g, w => w[0].toUpperCase() + w.slice(1).toLowerCase()),
  trim_:      ([t]) => t.trim().replace(/\s+/g, ' '),
  left:       ([t, n = 1]) => t.slice(0, +n),
  right:      ([t, n = 1]) => t.slice(-+n),
  mid:        ([t, s, n]) => t.slice(+s - 1, +s - 1 + +n),
  rept:       ([t, n]) => t.repeat(+n),
  exact:      ([a, b]) => a === b ? 'TRUE' : 'FALSE',
  find_:      ([t, s]) => { const i = s.indexOf(t); return i === -1 ? '#VALUE!' : i + 1; },
  substitute: ([t, old, n]) => t.split(old).join(n),
  value_:     ([t]) => isNaN(+t) ? '#VALUE!' : +t,
  code_:      ([t]) => t.charCodeAt(0),
  char_:      ([n]) => String.fromCharCode(+n),
  concat:     v => v.join(''),
};

window.FORMULAS = [

  // ═══════════════════════════════════════════════
  //  MATH & STATISTICS
  // ═══════════════════════════════════════════════
  {
    id: 'SUM', zh: '加總', desc: '計算一組數字的總和',
    cat: 'math', diff: 'b', ver: 'all',
    syntax: '=SUM(number1, [number2], ...)',
    args: [
      { n: 'number1', d: '要加總的第一個數字或範圍', req: true },
      { n: 'number2, ...', d: '額外的數字或範圍，最多 255 個', req: false }
    ],
    itype: 'nums', ivals: [10, 20, 30], ifn: 'sum',
    ex: '=SUM(10, 20, 30) → 60',
    tips: ['可加總整欄 =SUM(A:A)，忽略標題文字', '最多 255 個引數', '空格與文字自動忽略'],
    rel: ['SUMIF', 'SUMIFS', 'SUMPRODUCT']
  },
  {
    id: 'SUMIF', zh: '條件加總', desc: '加總符合單一條件的儲存格',
    cat: 'math', diff: 'i', ver: 'all',
    syntax: '=SUMIF(range, criteria, [sum_range])',
    args: [
      { n: 'range', d: '要評估條件的範圍', req: true },
      { n: 'criteria', d: '加總條件，如 ">100"、"蘋果"', req: true },
      { n: 'sum_range', d: '實際要加總的範圍（省略則加總 range）', req: false }
    ],
    itype: 'table',
    tableEx: {
      headers: ['品項', '金額'],
      rows: [['蘋果', 100], ['香蕉', 200], ['蘋果', 150], ['橘子', 80]],
      formula: '=SUMIF(A:A, "蘋果", B:B)',
      result: '250',
      note: '加總所有品項為「蘋果」的金額'
    },
    ex: '=SUMIF(A:A, "蘋果", B:B) → 250',
    tips: ['條件可用萬用字元：*（多字元）、?（單字元）', '數字條件用引號：">100"', '省略 sum_range 則加總 range 本身'],
    rel: ['SUMIFS', 'COUNTIF', 'AVERAGEIF']
  },
  {
    id: 'SUMIFS', zh: '多條件加總', desc: '加總同時符合多個條件的儲存格',
    cat: 'math', diff: 'i', ver: 'all',
    syntax: '=SUMIFS(sum_range, range1, criteria1, [range2, criteria2], ...)',
    args: [
      { n: 'sum_range', d: '要加總的範圍', req: true },
      { n: 'range1', d: '第一個條件範圍', req: true },
      { n: 'criteria1', d: '第一個條件', req: true },
      { n: 'range2, criteria2', d: '額外的條件範圍與條件（最多 127 組）', req: false }
    ],
    itype: 'none',
    ex: '=SUMIFS(C:C, A:A, "北區", B:B, "Q1") → 北區第一季總業績',
    tips: ['SUMIF 的多條件升級版', '所有範圍必須大小相同', '注意：SUMIFS 的 sum_range 在第一個位置（與 SUMIF 不同）'],
    rel: ['SUMIF', 'COUNTIFS', 'AVERAGEIFS']
  },
  {
    id: 'SUMPRODUCT', zh: '乘積加總', desc: '計算對應元素相乘後的總和，也可當多條件加總使用',
    cat: 'math', diff: 'a', ver: 'all',
    syntax: '=SUMPRODUCT(array1, [array2], ...)',
    args: [
      { n: 'array1', d: '第一個要相乘的陣列或範圍', req: true },
      { n: 'array2, ...', d: '其他要相乘的陣列（大小需相同）', req: false }
    ],
    itype: 'none',
    ex: '=SUMPRODUCT(B2:B5, C2:C5) → 各項目的數量×單價加總',
    tips: ['=SUMPRODUCT((A:A="蘋果")*(B:B>100)*C:C) 可取代 SUMIFS', '不需要 Ctrl+Shift+Enter，自動處理陣列', '效能比 SUMIFS 略慢但更靈活'],
    rel: ['SUMIFS', 'MMULT']
  },
  {
    id: 'AVERAGE', zh: '平均值', desc: '計算一組數字的算術平均值',
    cat: 'math', diff: 'b', ver: 'all',
    syntax: '=AVERAGE(number1, [number2], ...)',
    args: [
      { n: 'number1', d: '要計算平均的第一個數字或範圍', req: true },
      { n: 'number2, ...', d: '額外的數字或範圍', req: false }
    ],
    itype: 'nums', ivals: [80, 90, 70, 95], ifn: 'avg',
    ex: '=AVERAGE(80, 90, 70, 95) → 83.75',
    tips: ['忽略空白儲存格與文字', '若需含邏輯值請用 AVERAGEA', '與 AVERAGEIF 搭配可計算條件平均'],
    rel: ['AVERAGEIF', 'AVERAGEIFS', 'MEDIAN']
  },
  {
    id: 'AVERAGEIF', zh: '條件平均', desc: '計算符合單一條件的儲存格平均值',
    cat: 'math', diff: 'i', ver: 'all',
    syntax: '=AVERAGEIF(range, criteria, [average_range])',
    args: [
      { n: 'range', d: '要評估條件的範圍', req: true },
      { n: 'criteria', d: '平均條件', req: true },
      { n: 'average_range', d: '實際要平均的範圍', req: false }
    ],
    itype: 'none',
    ex: '=AVERAGEIF(A:A, "業務", B:B) → 業務部門的平均薪資',
    tips: ['條件支援萬用字元 * 和 ?', '省略 average_range 則平均 range 本身'],
    rel: ['AVERAGEIFS', 'SUMIF', 'COUNTIF']
  },
  {
    id: 'COUNT', zh: '計算數字個數', desc: '計算範圍內包含數字的儲存格數量',
    cat: 'math', diff: 'b', ver: 'all',
    syntax: '=COUNT(value1, [value2], ...)',
    args: [{ n: 'value1', d: '要計算的範圍', req: true }],
    itype: 'nums', ivals: [1, 2, 3], ifn: 'count',
    ex: '=COUNT(A1:A10) → 10（若有3格是文字則回傳7）',
    tips: ['只計算數字，忽略文字、空白、邏輯值', 'COUNTA 計算所有非空格儲存格', 'COUNTBLANK 計算空白儲存格'],
    rel: ['COUNTA', 'COUNTBLANK', 'COUNTIF']
  },
  {
    id: 'COUNTA', zh: '計算非空格數', desc: '計算範圍內所有非空白的儲存格數量',
    cat: 'math', diff: 'b', ver: 'all',
    syntax: '=COUNTA(value1, [value2], ...)',
    args: [{ n: 'value1', d: '要計算的範圍', req: true }],
    itype: 'none',
    ex: '=COUNTA(A1:A10) → 8（若有2格空白則回傳8）',
    tips: ['計算文字、數字、邏輯值、錯誤值', '不計算空白儲存格', '常用於動態計算清單長度'],
    rel: ['COUNT', 'COUNTBLANK', 'COUNTIF']
  },
  {
    id: 'COUNTIF', zh: '條件計算', desc: '計算符合單一條件的儲存格數量',
    cat: 'math', diff: 'i', ver: 'all',
    syntax: '=COUNTIF(range, criteria)',
    args: [
      { n: 'range', d: '要計算的範圍', req: true },
      { n: 'criteria', d: '計算條件', req: true }
    ],
    itype: 'none',
    ex: '=COUNTIF(A:A, ">100") → 計算大於100的儲存格數',
    tips: ['條件支援萬用字元 * 和 ?', '=COUNTIF(A:A, A2) 可找重複值', '不區分大小寫'],
    rel: ['COUNTIFS', 'SUMIF', 'COUNT']
  },
  {
    id: 'COUNTIFS', zh: '多條件計算', desc: '計算同時符合多個條件的儲存格數量',
    cat: 'math', diff: 'i', ver: 'all',
    syntax: '=COUNTIFS(range1, criteria1, [range2, criteria2], ...)',
    args: [
      { n: 'range1', d: '第一個條件範圍', req: true },
      { n: 'criteria1', d: '第一個條件', req: true },
      { n: 'range2, criteria2', d: '額外條件（最多127組）', req: false }
    ],
    itype: 'none',
    ex: '=COUNTIFS(A:A, "業務", B:B, ">50000") → 業務部門薪資超過5萬的人數',
    tips: ['所有條件範圍大小必須相同', '比 SUMPRODUCT 條件計算更快', '日期條件：=COUNTIFS(A:A,">="&DATE(2024,1,1))'],
    rel: ['COUNTIF', 'SUMIFS']
  },
  {
    id: 'MAX', zh: '最大值', desc: '回傳一組數字中的最大值',
    cat: 'math', diff: 'b', ver: 'all',
    syntax: '=MAX(number1, [number2], ...)',
    args: [{ n: 'number1', d: '要比較的數字或範圍', req: true }],
    itype: 'nums', ivals: [5, 12, 8, 3, 19], ifn: 'max',
    ex: '=MAX(5, 12, 8, 3, 19) → 19',
    tips: ['忽略文字與空白', 'MAXIFS（2019+）可加條件篩選最大值', '用 MAX-MIN 可計算全距'],
    rel: ['MIN', 'LARGE', 'MAXIFS']
  },
  {
    id: 'MIN', zh: '最小值', desc: '回傳一組數字中的最小值',
    cat: 'math', diff: 'b', ver: 'all',
    syntax: '=MIN(number1, [number2], ...)',
    args: [{ n: 'number1', d: '要比較的數字或範圍', req: true }],
    itype: 'nums', ivals: [5, 12, 8, 3, 19], ifn: 'min',
    ex: '=MIN(5, 12, 8, 3, 19) → 3',
    tips: ['忽略文字與空白', 'MINIFS（2019+）可加條件篩選最小值'],
    rel: ['MAX', 'SMALL', 'MINIFS']
  },
  {
    id: 'LARGE', zh: '第K大值', desc: '回傳資料集中第 K 大的數值',
    cat: 'math', diff: 'i', ver: 'all',
    syntax: '=LARGE(array, k)',
    args: [
      { n: 'array', d: '要排序的數值範圍', req: true },
      { n: 'k', d: '排名，1代表最大值', req: true }
    ],
    itype: 'none',
    ex: '=LARGE(A1:A10, 2) → 第二大的值',
    tips: ['k=1 等同於 MAX', '可用來找前三名：LARGE(A:A,1)、LARGE(A:A,2)、LARGE(A:A,3)'],
    rel: ['SMALL', 'MAX', 'RANK']
  },
  {
    id: 'SMALL', zh: '第K小值', desc: '回傳資料集中第 K 小的數值',
    cat: 'math', diff: 'i', ver: 'all',
    syntax: '=SMALL(array, k)',
    args: [
      { n: 'array', d: '要排序的數值範圍', req: true },
      { n: 'k', d: '排名，1代表最小值', req: true }
    ],
    itype: 'none',
    ex: '=SMALL(A1:A10, 3) → 第三小的值',
    tips: ['k=1 等同於 MIN', '常用於找後幾名或異常低值'],
    rel: ['LARGE', 'MIN', 'RANK']
  },
  {
    id: 'ROUND', zh: '四捨五入', desc: '將數字四捨五入到指定的小數位數',
    cat: 'math', diff: 'b', ver: 'all',
    syntax: '=ROUND(number, num_digits)',
    args: [
      { n: 'number', d: '要四捨五入的數字', req: true },
      { n: 'num_digits', d: '小數位數，負數代表整數位', req: true }
    ],
    itype: 'text2', ivals: ['3.14159', '2'], iparams: ['數字', '小數位'], ifn: 'round',
    ex: '=ROUND(3.14159, 2) → 3.14',
    tips: ['num_digits=0 四捨五入到整數', 'num_digits=-2 四捨五入到百位（如 1234→1200）', 'ROUNDUP 永遠無條件進位，ROUNDDOWN 永遠無條件捨去'],
    rel: ['ROUNDUP', 'ROUNDDOWN', 'INT', 'CEILING', 'FLOOR']
  },
  {
    id: 'ROUNDUP', zh: '無條件進位', desc: '將數字永遠向上捨入（無條件進位）',
    cat: 'math', diff: 'b', ver: 'all',
    syntax: '=ROUNDUP(number, num_digits)',
    args: [
      { n: 'number', d: '要進位的數字', req: true },
      { n: 'num_digits', d: '小數位數', req: true }
    ],
    itype: 'text2', ivals: ['2.1', '0'], iparams: ['數字', '小數位'], ifn: 'roundup',
    ex: '=ROUNDUP(2.1, 0) → 3',
    tips: ['無論小數多小都會進位', '用於計算需要「足夠」數量的情境，如包裝箱數'],
    rel: ['ROUNDDOWN', 'ROUND', 'CEILING']
  },
  {
    id: 'ROUNDDOWN', zh: '無條件捨去', desc: '將數字永遠向下捨入（無條件捨去）',
    cat: 'math', diff: 'b', ver: 'all',
    syntax: '=ROUNDDOWN(number, num_digits)',
    args: [
      { n: 'number', d: '要捨去的數字', req: true },
      { n: 'num_digits', d: '小數位數', req: true }
    ],
    itype: 'text2', ivals: ['2.9', '0'], iparams: ['數字', '小數位'], ifn: 'rounddown',
    ex: '=ROUNDDOWN(2.9, 0) → 2',
    tips: ['無論小數多大都會捨去', '計算員工完整年資常用'],
    rel: ['ROUNDUP', 'ROUND', 'INT', 'FLOOR']
  },
  {
    id: 'INT', zh: '取整數', desc: '將數字向下取至最接近的整數',
    cat: 'math', diff: 'b', ver: 'all',
    syntax: '=INT(number)',
    args: [{ n: 'number', d: '要取整的數字', req: true }],
    itype: 'text2', ivals: ['7.85', ''], iparams: ['數字', '（無第二參數）'], ifn: 'int_',
    ex: '=INT(7.85) → 7 ／ =INT(-7.85) → -8',
    tips: ['負數會向下（往更負的方向）取整，與 ROUNDDOWN 不同', '=A1-INT(A1) 可取得小數部分'],
    rel: ['ROUNDDOWN', 'TRUNC', 'MOD']
  },
  {
    id: 'MOD', zh: '取餘數', desc: '回傳除法的餘數',
    cat: 'math', diff: 'i', ver: 'all',
    syntax: '=MOD(number, divisor)',
    args: [
      { n: 'number', d: '被除數', req: true },
      { n: 'divisor', d: '除數', req: true }
    ],
    itype: 'text2', ivals: ['10', '3'], iparams: ['被除數', '除數'], ifn: 'mod',
    ex: '=MOD(10, 3) → 1',
    tips: ['=MOD(ROW(),2)=0 可標記偶數列', '用於判斷整除：MOD(A1,2)=0 代表偶數', '餘數的正負符號與除數相同'],
    rel: ['INT', 'QUOTIENT']
  },
  {
    id: 'ABS', zh: '絕對值', desc: '回傳數字的絕對值（去除正負號）',
    cat: 'math', diff: 'b', ver: 'all',
    syntax: '=ABS(number)',
    args: [{ n: 'number', d: '要取絕對值的數字', req: true }],
    itype: 'text2', ivals: ['-42', ''], iparams: ['數字', '（無第二參數）'], ifn: 'abs',
    ex: '=ABS(-42) → 42',
    tips: ['計算差距時常用：=ABS(A1-B1)', '配合 IF 可處理正負值'],
    rel: ['SIGN']
  },
  {
    id: 'POWER', zh: '次方', desc: '計算數字的指定次方',
    cat: 'math', diff: 'b', ver: 'all',
    syntax: '=POWER(number, power)',
    args: [
      { n: 'number', d: '底數', req: true },
      { n: 'power', d: '指數', req: true }
    ],
    itype: 'text2', ivals: ['2', '10'], iparams: ['底數', '指數'], ifn: 'power',
    ex: '=POWER(2, 10) → 1024',
    tips: ['也可用 ^ 運算符：=2^10', '=POWER(A1, 1/2) 等同 SQRT(A1)（平方根）'],
    rel: ['SQRT', 'EXP', 'LOG']
  },
  {
    id: 'SQRT', zh: '平方根', desc: '計算數字的正平方根',
    cat: 'math', diff: 'b', ver: 'all',
    syntax: '=SQRT(number)',
    args: [{ n: 'number', d: '要開根號的正數', req: true }],
    itype: 'text2', ivals: ['144', ''], iparams: ['數字', '（無第二參數）'], ifn: 'sqrt',
    ex: '=SQRT(144) → 12',
    tips: ['負數會回傳 #NUM! 錯誤', '用 =SQRT(ABS(A1)) 避免錯誤'],
    rel: ['POWER']
  },

  // ═══════════════════════════════════════════════
  //  LOOKUP & REFERENCE
  // ═══════════════════════════════════════════════
  {
    id: 'VLOOKUP', zh: '垂直查找', desc: '在表格第一欄查找值，回傳同列指定欄的資料',
    cat: 'lookup', diff: 'i', ver: 'all',
    syntax: '=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
    args: [
      { n: 'lookup_value', d: '要查找的值', req: true },
      { n: 'table_array', d: '查找表格範圍（第一欄為查找欄）', req: true },
      { n: 'col_index_num', d: '要回傳的欄號（從1開始）', req: true },
      { n: 'range_lookup', d: 'FALSE=精確比對（建議），TRUE=近似比對', req: false }
    ],
    itype: 'table',
    tableEx: {
      headers: ['員工ID', '姓名', '部門'],
      rows: [['A001', '王大明', '業務部'], ['A002', '李小華', '財務部'], ['A003', '張美玲', '人資部']],
      formula: '=VLOOKUP("A002", A:C, 2, FALSE)',
      result: '李小華',
      note: '查找ID "A002" 並回傳第2欄（姓名）'
    },
    ex: '=VLOOKUP("A002", A:C, 2, FALSE) → 李小華',
    tips: ['務必使用 FALSE 進行精確比對', '只能向右查找，無法向左', '考慮改用 XLOOKUP（365）或 INDEX+MATCH'],
    rel: ['XLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH']
  },
  {
    id: 'HLOOKUP', zh: '水平查找', desc: '在表格第一列查找值，回傳同欄指定列的資料',
    cat: 'lookup', diff: 'i', ver: 'all',
    syntax: '=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])',
    args: [
      { n: 'lookup_value', d: '要查找的值', req: true },
      { n: 'table_array', d: '查找表格（第一列為查找列）', req: true },
      { n: 'row_index_num', d: '要回傳的列號（從1開始）', req: true },
      { n: 'range_lookup', d: 'FALSE=精確比對', req: false }
    ],
    itype: 'none',
    ex: '=HLOOKUP("Q2", A1:D2, 2, FALSE) → 第二列對應Q2欄的值',
    tips: ['與 VLOOKUP 方向相反（橫向查找）', '適用於跨欄的季度/月份標題表', '現代替代方案：XLOOKUP'],
    rel: ['VLOOKUP', 'INDEX', 'MATCH', 'XLOOKUP']
  },
  {
    id: 'INDEX', zh: '索引取值', desc: '依據列號與欄號，從範圍中回傳特定位置的值',
    cat: 'lookup', diff: 'i', ver: 'all',
    syntax: '=INDEX(array, row_num, [col_num])',
    args: [
      { n: 'array', d: '查找範圍', req: true },
      { n: 'row_num', d: '要回傳的列號', req: true },
      { n: 'col_num', d: '要回傳的欄號（省略則回傳整列）', req: false }
    ],
    itype: 'none',
    ex: '=INDEX(A1:C5, 3, 2) → 第3列第2欄的值',
    tips: ['與 MATCH 搭配是 VLOOKUP 的強力替代方案', '可向左查找（VLOOKUP 做不到）', '省略列號或欄號時回傳整欄/整列'],
    rel: ['MATCH', 'VLOOKUP', 'OFFSET']
  },
  {
    id: 'MATCH', zh: '位置查找', desc: '在範圍中查找值，回傳其相對位置（數字）',
    cat: 'lookup', diff: 'i', ver: 'all',
    syntax: '=MATCH(lookup_value, lookup_array, [match_type])',
    args: [
      { n: 'lookup_value', d: '要查找的值', req: true },
      { n: 'lookup_array', d: '查找範圍（單列或單欄）', req: true },
      { n: 'match_type', d: '0=精確比對，1=最大小於，-1=最小大於', req: false }
    ],
    itype: 'none',
    ex: '=MATCH("李小華", A:A, 0) → 3（在第3列找到）',
    tips: ['通常與 INDEX 搭配使用', 'match_type=0 為精確比對（最常用）', '找不到時回傳 #N/A'],
    rel: ['INDEX', 'XMATCH', 'VLOOKUP']
  },
  {
    id: 'CHOOSE', zh: '依編號選取', desc: '根據索引號從清單中選取對應的值',
    cat: 'lookup', diff: 'i', ver: 'all',
    syntax: '=CHOOSE(index_num, value1, [value2], ...)',
    args: [
      { n: 'index_num', d: '要選取的位置（1=第一個）', req: true },
      { n: 'value1, ...', d: '可選取的值清單', req: true }
    ],
    itype: 'none',
    ex: '=CHOOSE(2, "一月", "二月", "三月") → 二月',
    tips: ['最多 254 個選項', '可用 WEEKDAY 搭配顯示星期名稱', '=CHOOSE(MONTH(TODAY()),\"Q1\",\"Q1\",\"Q1\",\"Q2\",...) 轉換季度'],
    rel: ['IF', 'IFS', 'SWITCH']
  },
  {
    id: 'OFFSET', zh: '偏移參照', desc: '從基準儲存格偏移指定列欄後，回傳範圍參照',
    cat: 'lookup', diff: 'a', ver: 'all',
    syntax: '=OFFSET(reference, rows, cols, [height], [width])',
    args: [
      { n: 'reference', d: '起始儲存格', req: true },
      { n: 'rows', d: '向下偏移的列數（負數向上）', req: true },
      { n: 'cols', d: '向右偏移的欄數（負數向左）', req: true },
      { n: 'height', d: '回傳範圍的高度', req: false },
      { n: 'width', d: '回傳範圍的寬度', req: false }
    ],
    itype: 'none',
    ex: '=OFFSET(A1, 2, 1) → C3儲存格的值',
    tips: ['為「揮發性函數」，每次計算都會重新計算（效能較差）', '常用於動態命名範圍', '配合 COUNTA 建立動態下拉清單'],
    rel: ['INDEX', 'INDIRECT']
  },
  {
    id: 'INDIRECT', zh: '間接參照', desc: '將文字字串轉換為有效的儲存格參照',
    cat: 'lookup', diff: 'a', ver: 'all',
    syntax: '=INDIRECT(ref_text, [a1])',
    args: [
      { n: 'ref_text', d: '要轉換為參照的文字字串', req: true },
      { n: 'a1', d: 'TRUE=A1樣式，FALSE=R1C1樣式', req: false }
    ],
    itype: 'none',
    ex: '=INDIRECT("A"&B1) → 依 B1 的值動態指向不同列',
    tips: ['為「揮發性函數」，效能較差', '常用於跨工作表動態參照', '=SUM(INDIRECT(A1&"!B:B")) 動態指定工作表'],
    rel: ['OFFSET', 'ADDRESS', 'INDEX']
  },
  {
    id: 'ROW', zh: '列號', desc: '回傳儲存格的列號',
    cat: 'lookup', diff: 'b', ver: 'all',
    syntax: '=ROW([reference])',
    args: [{ n: 'reference', d: '要取得列號的儲存格（省略則回傳公式所在列）', req: false }],
    itype: 'none',
    ex: '=ROW(A5) → 5 ／ =ROW() → 目前儲存格的列號',
    tips: ['=ROW(A1)-1 建立從0開始的序號', 'MOD(ROW(),2) 常用於斑馬紋格式化', '=ROW(A1:A10) 作為陣列公式產生1到10的序列'],
    rel: ['COLUMN', 'ROWS', 'SEQUENCE']
  },
  {
    id: 'COLUMN', zh: '欄號', desc: '回傳儲存格的欄號',
    cat: 'lookup', diff: 'b', ver: 'all',
    syntax: '=COLUMN([reference])',
    args: [{ n: 'reference', d: '要取得欄號的儲存格', req: false }],
    itype: 'none',
    ex: '=COLUMN(C1) → 3',
    tips: ['A欄=1、B欄=2、以此類推', '=ADDRESS(1,COLUMN()) 取得欄位字母'],
    rel: ['ROW', 'COLUMNS', 'ADDRESS']
  },

  // ═══════════════════════════════════════════════
  //  TEXT
  // ═══════════════════════════════════════════════
  {
    id: 'LEFT', zh: '左取文字', desc: '從文字左側取出指定數量的字元',
    cat: 'text', diff: 'b', ver: 'all',
    syntax: '=LEFT(text, [num_chars])',
    args: [
      { n: 'text', d: '來源文字', req: true },
      { n: 'num_chars', d: '要取出的字元數（預設1）', req: false }
    ],
    itype: 'text2', ivals: ['Excel公式大師', '5'], iparams: ['文字', '字元數'], ifn: 'left',
    ex: '=LEFT("Excel公式大師", 5) → Excel',
    tips: ['中文字元也算1個字元', '常用於提取代碼前綴、電話區碼等'],
    rel: ['RIGHT', 'MID', 'LEN', 'FIND']
  },
  {
    id: 'RIGHT', zh: '右取文字', desc: '從文字右側取出指定數量的字元',
    cat: 'text', diff: 'b', ver: 'all',
    syntax: '=RIGHT(text, [num_chars])',
    args: [
      { n: 'text', d: '來源文字', req: true },
      { n: 'num_chars', d: '要取出的字元數（預設1）', req: false }
    ],
    itype: 'text2', ivals: ['PROD-2024-Q3', '2'], iparams: ['文字', '字元數'], ifn: 'right',
    ex: '=RIGHT("PROD-2024-Q3", 2) → Q3',
    tips: ['常用於提取尾碼、年份末兩碼等'],
    rel: ['LEFT', 'MID', 'LEN']
  },
  {
    id: 'MID', zh: '中間取字', desc: '從文字的指定位置開始，取出指定數量的字元',
    cat: 'text', diff: 'b', ver: 'all',
    syntax: '=MID(text, start_num, num_chars)',
    args: [
      { n: 'text', d: '來源文字', req: true },
      { n: 'start_num', d: '起始位置（從1開始）', req: true },
      { n: 'num_chars', d: '要取出的字元數', req: true }
    ],
    itype: 'text3', ivals: ['A001-王大明-業務', '6', '3'], iparams: ['文字', '起始位置', '字元數'], ifn: 'mid',
    ex: '=MID("A001-王大明-業務", 6, 3) → 王大明',
    tips: ['搭配 FIND 可動態定位起始位置', '=MID(A1, FIND("-",A1)+1, 99) 取出第一個「-」後的所有文字'],
    rel: ['LEFT', 'RIGHT', 'FIND', 'LEN']
  },
  {
    id: 'LEN', zh: '字元長度', desc: '回傳文字字串的字元數',
    cat: 'text', diff: 'b', ver: 'all',
    syntax: '=LEN(text)',
    args: [{ n: 'text', d: '要計算長度的文字', req: true }],
    itype: 'text1', ival: 'Excel公式動態圖卡', ifn: 'len',
    ex: '=LEN("Excel公式動態圖卡") → 10',
    tips: ['空格也算1個字元', '常用於驗證身分證、電話格式', '=LEN(A1)-LEN(SUBSTITUTE(A1,"-","")) 計算特定字元出現次數'],
    rel: ['LEFT', 'RIGHT', 'MID', 'TRIM']
  },
  {
    id: 'FIND', zh: '尋找位置（區分大小寫）', desc: '在文字中尋找子字串，回傳起始位置（區分大小寫）',
    cat: 'text', diff: 'i', ver: 'all',
    syntax: '=FIND(find_text, within_text, [start_num])',
    args: [
      { n: 'find_text', d: '要尋找的文字', req: true },
      { n: 'within_text', d: '要搜尋的來源文字', req: true },
      { n: 'start_num', d: '開始搜尋的位置（預設1）', req: false }
    ],
    itype: 'text2', ivals: ['-', 'PROD-2024-Q3'], iparams: ['搜尋字元', '來源文字'], ifn: 'find_',
    ex: '=FIND("-", "PROD-2024-Q3") → 5',
    tips: ['區分大小寫，找不到回傳 #VALUE!', 'SEARCH 不區分大小寫且支援萬用字元', '用 IFERROR 包住避免錯誤值顯示'],
    rel: ['SEARCH', 'SUBSTITUTE', 'MID']
  },
  {
    id: 'SEARCH', zh: '搜尋位置（不區分大小寫）', desc: '在文字中搜尋子字串，不區分大小寫，支援萬用字元',
    cat: 'text', diff: 'i', ver: 'all',
    syntax: '=SEARCH(find_text, within_text, [start_num])',
    args: [
      { n: 'find_text', d: '要搜尋的文字（支援*和?萬用字元）', req: true },
      { n: 'within_text', d: '要搜尋的來源文字', req: true },
      { n: 'start_num', d: '開始搜尋的位置', req: false }
    ],
    itype: 'none',
    ex: '=SEARCH("excel", "Microsoft Excel") → 11（不區分大小寫）',
    tips: ['不區分大小寫，與 FIND 的關鍵差異', '支援萬用字元：*代表任意字元，?代表單一字元', '找不到回傳 #VALUE!'],
    rel: ['FIND', 'SUBSTITUTE']
  },
  {
    id: 'SUBSTITUTE', zh: '替換文字', desc: '將文字中所有（或指定次）舊文字替換為新文字',
    cat: 'text', diff: 'i', ver: 'all',
    syntax: '=SUBSTITUTE(text, old_text, new_text, [instance_num])',
    args: [
      { n: 'text', d: '來源文字', req: true },
      { n: 'old_text', d: '要被替換的文字', req: true },
      { n: 'new_text', d: '替換後的新文字', req: true },
      { n: 'instance_num', d: '指定替換第幾次出現（省略則全部替換）', req: false }
    ],
    itype: 'text3', ivals: ['2024-01-15', '-', '/'], iparams: ['來源', '舊字元', '新字元'], ifn: 'substitute',
    ex: '=SUBSTITUTE("2024-01-15", "-", "/") → 2024/01/15',
    tips: ['區分大小寫', '=SUBSTITUTE(A1," ","") 刪除所有空格', '=LEN(A1)-LEN(SUBSTITUTE(A1,"@","")) 計算@出現次數'],
    rel: ['REPLACE', 'FIND', 'TRIM']
  },
  {
    id: 'REPLACE', zh: '位置替換', desc: '依據起始位置與字元數，替換文字的特定部分',
    cat: 'text', diff: 'i', ver: 'all',
    syntax: '=REPLACE(old_text, start_num, num_chars, new_text)',
    args: [
      { n: 'old_text', d: '來源文字', req: true },
      { n: 'start_num', d: '替換的起始位置', req: true },
      { n: 'num_chars', d: '要替換的字元數', req: true },
      { n: 'new_text', d: '替換後的新文字', req: true }
    ],
    itype: 'none',
    ex: '=REPLACE("A001-明", 1, 4, "B999") → B999-明',
    tips: ['用於固定位置的替換，SUBSTITUTE 用於特定內容的替換', '設 new_text="" 可刪除指定部分'],
    rel: ['SUBSTITUTE', 'MID', 'LEFT']
  },
  {
    id: 'UPPER', zh: '轉大寫', desc: '將文字中的所有英文字母轉換為大寫',
    cat: 'text', diff: 'b', ver: 'all',
    syntax: '=UPPER(text)',
    args: [{ n: 'text', d: '要轉換的文字', req: true }],
    itype: 'text1', ival: 'hello world', ifn: 'upper',
    ex: '=UPPER("hello world") → HELLO WORLD',
    tips: ['非英文字母不受影響', '與 EXACT 搭配可做大小寫敏感比對'],
    rel: ['LOWER', 'PROPER']
  },
  {
    id: 'LOWER', zh: '轉小寫', desc: '將文字中的所有英文字母轉換為小寫',
    cat: 'text', diff: 'b', ver: 'all',
    syntax: '=LOWER(text)',
    args: [{ n: 'text', d: '要轉換的文字', req: true }],
    itype: 'text1', ival: 'EXCEL MASTER', ifn: 'lower',
    ex: '=LOWER("EXCEL MASTER") → excel master',
    tips: ['常用於標準化電子郵件地址', '配合 TRIM 清理資料'],
    rel: ['UPPER', 'PROPER']
  },
  {
    id: 'PROPER', zh: '首字大寫', desc: '將每個英文單詞的首字母轉為大寫，其餘轉小寫',
    cat: 'text', diff: 'b', ver: 'all',
    syntax: '=PROPER(text)',
    args: [{ n: 'text', d: '要轉換的文字', req: true }],
    itype: 'text1', ival: 'john smith - manager', ifn: 'proper',
    ex: '=PROPER("john smith - manager") → John Smith - Manager',
    tips: ['每個空格、連字符、標點後的字母都會大寫', '中文不受影響'],
    rel: ['UPPER', 'LOWER']
  },
  {
    id: 'TRIM', zh: '去除多餘空格', desc: '刪除文字前後的空格，並將中間多個連續空格縮減為一個',
    cat: 'text', diff: 'b', ver: 'all',
    syntax: '=TRIM(text)',
    args: [{ n: 'text', d: '要清理空格的文字', req: true }],
    itype: 'text1', ival: '  Excel  公式  ', ifn: 'trim_',
    ex: '=TRIM("  Excel  公式  ") → Excel 公式',
    tips: ['只去除 ASCII 空格（字元32），不去除不換行空格（CHAR(160)）', '資料清理必備，處理貼上的雜亂資料', '搭配 CLEAN 一起使用效果更好'],
    rel: ['CLEAN', 'SUBSTITUTE', 'LEN']
  },
  {
    id: 'TEXT', zh: '格式化文字', desc: '將數字或日期以指定格式轉換為文字',
    cat: 'text', diff: 'i', ver: 'all',
    syntax: '=TEXT(value, format_text)',
    args: [
      { n: 'value', d: '要格式化的數字或日期', req: true },
      { n: 'format_text', d: '格式代碼（用引號括起）', req: true }
    ],
    itype: 'none',
    ex: '=TEXT(1234567.89, "#,##0.00") → 1,234,567.89\n=TEXT(TODAY(), "yyyy年mm月dd日") → 2024年01月15日',
    tips: ['格式碼：0=補零，#=不補零，,=千分位，.=小數點', '日期格式：yyyy=4位年，mm=2位月，dd=2位日', '轉換後為文字，無法再做數學運算'],
    rel: ['VALUE', 'CONCATENATE', 'TODAY']
  },
  {
    id: 'VALUE', zh: '文字轉數字', desc: '將代表數字的文字字串轉換為數值',
    cat: 'text', diff: 'b', ver: 'all',
    syntax: '=VALUE(text)',
    args: [{ n: 'text', d: '要轉換的文字（必須是有效數字格式）', req: true }],
    itype: 'text1', ival: '42', ifn: 'value_',
    ex: '=VALUE("42") → 42（數字）',
    tips: ['從系統匯入的資料常有「數字存成文字」的問題', '也可直接用 --A1 或 A1*1 強制轉換', '無效文字回傳 #VALUE!'],
    rel: ['TEXT', 'NUMBERVALUE']
  },
  {
    id: 'CONCATENATE', zh: '連接文字', desc: '將多個文字字串合併成一個',
    cat: 'text', diff: 'b', ver: 'all',
    syntax: '=CONCATENATE(text1, [text2], ...)',
    args: [{ n: 'text1, text2, ...', d: '要合併的文字或儲存格', req: true }],
    itype: 'none',
    ex: '=CONCATENATE("姓：",A1,"，名：",B1) → 姓：王，名：大明',
    tips: ['現代 Excel 建議改用 & 運算符：=A1&"，"&B1', 'CONCAT（2016+）和 TEXTJOIN（2016+）更強大', '最多 255 個引數'],
    rel: ['CONCAT', 'TEXTJOIN', 'TEXT']
  },
  {
    id: 'TEXTJOIN', zh: '分隔符連接', desc: '用指定分隔符連接多個文字，可忽略空白格',
    cat: 'text', diff: 'i', ver: '2016',
    syntax: '=TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)',
    args: [
      { n: 'delimiter', d: '分隔符（如逗號、空格）', req: true },
      { n: 'ignore_empty', d: 'TRUE=忽略空白格，FALSE=包含空白格', req: true },
      { n: 'text1, ...', d: '要連接的文字或範圍', req: true }
    ],
    itype: 'none',
    ex: '=TEXTJOIN(", ", TRUE, A1:A5) → 蘋果, 香蕉, 橘子（忽略空格）',
    tips: ['可直接指定範圍，比 CONCATENATE 省力很多', '搭配 IF 可實現條件連接：=TEXTJOIN(", ",TRUE,IF(B:B="業務",A:A,""))'],
    rel: ['CONCAT', 'CONCATENATE']
  },
  {
    id: 'EXACT', zh: '完全比對', desc: '比對兩個文字是否完全相同（區分大小寫）',
    cat: 'text', diff: 'i', ver: 'all',
    syntax: '=EXACT(text1, text2)',
    args: [
      { n: 'text1', d: '第一個文字', req: true },
      { n: 'text2', d: '第二個文字', req: true }
    ],
    itype: 'text2', ivals: ['Excel', 'excel'], iparams: ['文字1', '文字2'], ifn: 'exact',
    ex: '=EXACT("Excel", "excel") → FALSE',
    tips: ['= 運算符不區分大小寫，EXACT 才會', '用於密碼、代碼等需要嚴格大小寫比對的場景'],
    rel: ['FIND', 'IF']
  },
  {
    id: 'REPT', zh: '重複文字', desc: '將文字重複指定次數',
    cat: 'text', diff: 'b', ver: 'all',
    syntax: '=REPT(text, number_times)',
    args: [
      { n: 'text', d: '要重複的文字', req: true },
      { n: 'number_times', d: '重複次數', req: true }
    ],
    itype: 'text2', ivals: ['★', '5'], iparams: ['文字', '次數'], ifn: 'rept',
    ex: '=REPT("★", 5) → ★★★★★',
    tips: ['可用於製作星星評分、進度條視覺效果', '=REPT("█",A1/MAX(A:A)*10) 製作橫條圖'],
    rel: ['LEN', 'TEXT']
  },

  // ═══════════════════════════════════════════════
  //  LOGIC
  // ═══════════════════════════════════════════════
  {
    id: 'IF', zh: '條件判斷', desc: '若條件為真則回傳一個值，否則回傳另一個值',
    cat: 'logic', diff: 'b', ver: 'all',
    syntax: '=IF(logical_test, value_if_true, [value_if_false])',
    args: [
      { n: 'logical_test', d: '要評估的條件（回傳 TRUE 或 FALSE）', req: true },
      { n: 'value_if_true', d: '條件為真時回傳的值', req: true },
      { n: 'value_if_false', d: '條件為假時回傳的值（省略則回傳 FALSE）', req: false }
    ],
    itype: 'none',
    ex: '=IF(A1>=60, "及格", "不及格")',
    tips: ['可以最多巢狀 64 層（但超過 3 層建議改用 IFS）', '=IF(A1="","", ...) 避免空白格顯示錯誤', '巢狀 IF 可建立多段式評分'],
    rel: ['IFS', 'AND', 'OR', 'SWITCH', 'IFERROR']
  },
  {
    id: 'IFS', zh: '多條件判斷', desc: '依序測試多個條件，回傳第一個為真的條件對應值',
    cat: 'logic', diff: 'i', ver: '2016',
    syntax: '=IFS(logical_test1, value1, [logical_test2, value2], ...)',
    args: [
      { n: 'logical_test1', d: '第一個條件', req: true },
      { n: 'value1', d: '條件1為真時的回傳值', req: true },
      { n: '...', d: '更多條件與回傳值（最多 127 組）', req: false }
    ],
    itype: 'none',
    ex: '=IFS(A1>=90,"A",A1>=80,"B",A1>=70,"C",TRUE,"D")',
    tips: ['用 TRUE 作為最後條件，等同於 ELSE 的效果', '比巢狀 IF 更容易閱讀和維護', '若所有條件都不符合且無 TRUE，回傳 #N/A'],
    rel: ['IF', 'SWITCH', 'CHOOSE']
  },
  {
    id: 'AND', zh: '且（全部條件）', desc: '當所有條件都為真時回傳 TRUE，否則回傳 FALSE',
    cat: 'logic', diff: 'b', ver: 'all',
    syntax: '=AND(logical1, [logical2], ...)',
    args: [
      { n: 'logical1', d: '第一個要測試的條件', req: true },
      { n: 'logical2, ...', d: '額外的條件（最多255個）', req: false }
    ],
    itype: 'none',
    ex: '=AND(A1>60, B1="通過") → 兩個條件都要成立',
    tips: ['常與 IF 搭配：=IF(AND(...),"符合","不符合")', '所有條件都為 TRUE 才回傳 TRUE'],
    rel: ['OR', 'NOT', 'IF']
  },
  {
    id: 'OR', zh: '或（任一條件）', desc: '當任一條件為真時回傳 TRUE，全部為假才回傳 FALSE',
    cat: 'logic', diff: 'b', ver: 'all',
    syntax: '=OR(logical1, [logical2], ...)',
    args: [
      { n: 'logical1', d: '第一個要測試的條件', req: true },
      { n: 'logical2, ...', d: '額外的條件', req: false }
    ],
    itype: 'none',
    ex: '=OR(A1="業務", A1="行銷") → 業務或行銷部門都符合',
    tips: ['常與 IF 搭配', '只要有一個條件為 TRUE 即回傳 TRUE'],
    rel: ['AND', 'NOT', 'IF']
  },
  {
    id: 'NOT', zh: '否定', desc: '反轉邏輯值：TRUE 變 FALSE，FALSE 變 TRUE',
    cat: 'logic', diff: 'b', ver: 'all',
    syntax: '=NOT(logical)',
    args: [{ n: 'logical', d: '要反轉的邏輯值或條件', req: true }],
    itype: 'none',
    ex: '=NOT(A1>100) → A1不大於100時為TRUE',
    tips: ['=IF(NOT(ISBLANK(A1)), ...) 常見用法', '有時直接改寫條件更清晰，如 NOT(A>B) 等同 A<=B'],
    rel: ['AND', 'OR', 'IF']
  },
  {
    id: 'IFERROR', zh: '錯誤處理', desc: '若公式回傳錯誤，改回傳指定的替代值',
    cat: 'logic', diff: 'i', ver: 'all',
    syntax: '=IFERROR(value, value_if_error)',
    args: [
      { n: 'value', d: '要計算的公式或值', req: true },
      { n: 'value_if_error', d: '發生錯誤時的替代值', req: true }
    ],
    itype: 'none',
    ex: '=IFERROR(VLOOKUP(A1, B:C, 2, 0), "找不到") → 查無資料時顯示"找不到"',
    tips: ['適用所有錯誤：#N/A、#VALUE!、#REF!、#DIV/0!、#NUM!、#NAME?', '替代值可以是空字串 ""（不顯示任何內容）', 'IFNA 只捕捉 #N/A 錯誤，更精確'],
    rel: ['IFNA', 'ISERROR', 'IF']
  },
  {
    id: 'IFNA', zh: '#N/A 錯誤處理', desc: '若公式回傳 #N/A 錯誤，改回傳指定的替代值',
    cat: 'logic', diff: 'i', ver: 'all',
    syntax: '=IFNA(value, value_if_na)',
    args: [
      { n: 'value', d: '要計算的公式', req: true },
      { n: 'value_if_na', d: '回傳 #N/A 時的替代值', req: true }
    ],
    itype: 'none',
    ex: '=IFNA(VLOOKUP(A1,B:C,2,0), "新客戶") → 查無時顯示"新客戶"',
    tips: ['比 IFERROR 更精準，只處理 #N/A（查無結果）', '其他錯誤如 #VALUE! 不會被捕捉，有助於除錯'],
    rel: ['IFERROR', 'VLOOKUP', 'XLOOKUP']
  },
  {
    id: 'SWITCH', zh: '多值切換', desc: '將一個值與多個可能值比對，回傳第一個相符的結果',
    cat: 'logic', diff: 'i', ver: '2016',
    syntax: '=SWITCH(expression, value1, result1, [value2, result2], ..., [default])',
    args: [
      { n: 'expression', d: '要比對的值', req: true },
      { n: 'value1, result1', d: '比對值與對應結果', req: true },
      { n: '...default', d: '最後可加預設值（無符合時使用）', req: false }
    ],
    itype: 'none',
    ex: '=SWITCH(WEEKDAY(TODAY()),1,"週日",2,"週一",7,"週六","平日")',
    tips: ['比多層 IF 更簡潔（用於精確值比對，非範圍比對）', '最後加一個值作為預設（等同 ELSE）', '範圍比對（如 >=60）請用 IFS'],
    rel: ['IFS', 'IF', 'CHOOSE']
  },

  // ═══════════════════════════════════════════════
  //  DATE & TIME
  // ═══════════════════════════════════════════════
  {
    id: 'TODAY', zh: '今天日期', desc: '回傳今天的日期（每次開啟檔案或重新計算時更新）',
    cat: 'datetime', diff: 'b', ver: 'all',
    syntax: '=TODAY()',
    args: [],
    itype: 'date', dateEx: () => new Date().toLocaleDateString('zh-TW'),
    ex: '=TODAY() → 今天的日期（動態更新）',
    tips: ['揮發性函數，每次計算都會更新', '=TODAY()-A1 計算距今天數', '=TEXT(TODAY(),"yyyy/mm/dd") 格式化顯示'],
    rel: ['NOW', 'DATE', 'DATEDIF', 'TEXT']
  },
  {
    id: 'NOW', zh: '現在時間', desc: '回傳目前的日期與時間（揮發性，每次計算更新）',
    cat: 'datetime', diff: 'b', ver: 'all',
    syntax: '=NOW()',
    args: [],
    itype: 'date', dateEx: () => new Date().toLocaleString('zh-TW'),
    ex: '=NOW() → 目前日期和時間',
    tips: ['每次計算都會更新（揮發性函數）', '若只需要時間部分：=NOW()-TODAY()', '若需要固定時間戳記，請用 Ctrl+; 或 Ctrl+Shift+;'],
    rel: ['TODAY', 'HOUR', 'MINUTE', 'SECOND']
  },
  {
    id: 'DATE', zh: '建立日期', desc: '用年、月、日三個數字建立一個日期值',
    cat: 'datetime', diff: 'b', ver: 'all',
    syntax: '=DATE(year, month, day)',
    args: [
      { n: 'year', d: '年份（建議使用4位數）', req: true },
      { n: 'month', d: '月份（1-12，超過會自動跨年）', req: true },
      { n: 'day', d: '日期（超過月份天數會自動跨月）', req: true }
    ],
    itype: 'none',
    ex: '=DATE(2024, 1, 15) → 2024/1/15\n=DATE(2024, 13, 1) → 2025/1/1（自動跨年）',
    tips: ['用儲存格參照動態建立日期：=DATE(YEAR(A1),MONTH(A1)+1,1) 取下個月1日', '月、日超出範圍時自動進位，非常實用'],
    rel: ['YEAR', 'MONTH', 'DAY', 'TODAY', 'EDATE']
  },
  {
    id: 'YEAR', zh: '取出年份', desc: '從日期中取出年份',
    cat: 'datetime', diff: 'b', ver: 'all',
    syntax: '=YEAR(serial_number)',
    args: [{ n: 'serial_number', d: '日期值或儲存格', req: true }],
    itype: 'date',
    ex: '=YEAR(TODAY()) → 今年年份（如 2024）',
    tips: ['=YEAR(TODAY()) 取得目前年份', '與 DATE 搭配可以操作特定年份的日期'],
    rel: ['MONTH', 'DAY', 'DATE', 'TEXT']
  },
  {
    id: 'MONTH', zh: '取出月份', desc: '從日期中取出月份（1-12）',
    cat: 'datetime', diff: 'b', ver: 'all',
    syntax: '=MONTH(serial_number)',
    args: [{ n: 'serial_number', d: '日期值或儲存格', req: true }],
    itype: 'date',
    ex: '=MONTH(TODAY()) → 今天的月份（1-12）',
    tips: ['=CHOOSE(MONTH(A1),"一月","二月",...) 轉換為中文月份', '=MONTH(TODAY())>=7 判斷是否下半年'],
    rel: ['YEAR', 'DAY', 'DATE']
  },
  {
    id: 'DAY', zh: '取出日期', desc: '從日期中取出日（1-31）',
    cat: 'datetime', diff: 'b', ver: 'all',
    syntax: '=DAY(serial_number)',
    args: [{ n: 'serial_number', d: '日期值或儲存格', req: true }],
    itype: 'date',
    ex: '=DAY(TODAY()) → 今天是幾號',
    tips: ['=DAY(EOMONTH(TODAY(),0)) 取得本月最後一天是幾號', '=DATE(YEAR(A1),MONTH(A1)+1,0) 取得本月最後一天'],
    rel: ['YEAR', 'MONTH', 'EOMONTH']
  },
  {
    id: 'WEEKDAY', zh: '星期幾', desc: '回傳日期對應的星期數字（預設 1=週日，7=週六）',
    cat: 'datetime', diff: 'i', ver: 'all',
    syntax: '=WEEKDAY(serial_number, [return_type])',
    args: [
      { n: 'serial_number', d: '日期', req: true },
      { n: 'return_type', d: '1=週日起、2=週一起、3=週一起從0開始', req: false }
    ],
    itype: 'none',
    ex: '=WEEKDAY(TODAY(), 2) → 今天星期幾（1=週一，7=週日）',
    tips: ['return_type=2 時 1=週一，符合台灣習慣', '=TEXT(TODAY(),"dddd") 直接取得英文星期名', '=CHOOSE(WEEKDAY(TODAY(),2),"一","二","三","四","五","六","日") 取中文星期'],
    rel: ['TODAY', 'CHOOSE', 'NETWORKDAYS']
  },
  {
    id: 'DATEDIF', zh: '日期差距', desc: '計算兩個日期之間的差距（年、月或日）',
    cat: 'datetime', diff: 'i', ver: 'all',
    syntax: '=DATEDIF(start_date, end_date, unit)',
    args: [
      { n: 'start_date', d: '起始日期', req: true },
      { n: 'end_date', d: '結束日期（必須晚於起始日期）', req: true },
      { n: 'unit', d: '"Y"=年數，"M"=月數，"D"=天數，"YM"=扣除年的月，"MD"=扣除月的天', req: true }
    ],
    itype: 'none',
    ex: '=DATEDIF(A1, TODAY(), "Y") → 到今天為止的完整年數（年資）',
    tips: ['此函數在 Excel 說明中刻意隱藏，但完全有效', '計算年齡：=DATEDIF(生日, TODAY(), "Y")', '精確年資：=DATEDIF(入職日, TODAY(),"Y")&"年"&DATEDIF(入職日,TODAY(),"YM")&"個月"'],
    rel: ['TODAY', 'DATE', 'EDATE', 'NETWORKDAYS']
  },
  {
    id: 'EDATE', zh: '月份偏移', desc: '回傳從指定日期往前或往後偏移指定月數的日期',
    cat: 'datetime', diff: 'i', ver: 'all',
    syntax: '=EDATE(start_date, months)',
    args: [
      { n: 'start_date', d: '起始日期', req: true },
      { n: 'months', d: '偏移月數（負數往前）', req: true }
    ],
    itype: 'none',
    ex: '=EDATE(TODAY(), 3) → 三個月後的日期\n=EDATE(TODAY(), -1) → 一個月前的日期',
    tips: ['自動處理月底日期（如1月31日+1月→2月28日）', '用於計算合約到期日、訂閱續約日'],
    rel: ['EOMONTH', 'DATE', 'DATEDIF']
  },
  {
    id: 'EOMONTH', zh: '月底日期', desc: '回傳偏移指定月數後的月份最後一天',
    cat: 'datetime', diff: 'i', ver: 'all',
    syntax: '=EOMONTH(start_date, months)',
    args: [
      { n: 'start_date', d: '起始日期', req: true },
      { n: 'months', d: '偏移月數（0=本月，-1=上月，1=下月）', req: true }
    ],
    itype: 'none',
    ex: '=EOMONTH(TODAY(), 0) → 本月最後一天\n=EOMONTH(TODAY(), 1)+1 → 下個月第一天',
    tips: ['=EOMONTH(A1, 0)-DAY(A1)+1 取得本月第一天', '自動計算閏年（2月）'],
    rel: ['EDATE', 'DAY', 'DATE']
  },
  {
    id: 'NETWORKDAYS', zh: '工作天數', desc: '計算兩個日期之間的工作天數（排除週末，可自訂假日）',
    cat: 'datetime', diff: 'i', ver: 'all',
    syntax: '=NETWORKDAYS(start_date, end_date, [holidays])',
    args: [
      { n: 'start_date', d: '起始日期', req: true },
      { n: 'end_date', d: '結束日期', req: true },
      { n: 'holidays', d: '要排除的假日清單範圍', req: false }
    ],
    itype: 'none',
    ex: '=NETWORKDAYS(A1, B1, H:H) → 不含假日的工作天數',
    tips: ['包含起始和結束日期', 'NETWORKDAYS.INTL 可自訂週末（適合非標準工作週）', '用於計算交貨期、SLA 天數'],
    rel: ['WORKDAY', 'WEEKDAY', 'DATEDIF']
  },
  {
    id: 'WORKDAY', zh: '工作日偏移', desc: '從指定日期往後偏移指定工作天後的日期',
    cat: 'datetime', diff: 'i', ver: 'all',
    syntax: '=WORKDAY(start_date, days, [holidays])',
    args: [
      { n: 'start_date', d: '起始日期', req: true },
      { n: 'days', d: '要偏移的工作天數', req: true },
      { n: 'holidays', d: '額外排除的假日清單', req: false }
    ],
    itype: 'none',
    ex: '=WORKDAY(TODAY(), 5) → 5個工作日後的日期',
    tips: ['自動跳過週六、週日', '可傳入假日範圍排除特定日期', 'WORKDAY.INTL 可自訂週末'],
    rel: ['NETWORKDAYS', 'EDATE']
  },

  // ═══════════════════════════════════════════════
  //  ADVANCED — Excel 365 / 2019
  // ═══════════════════════════════════════════════
  {
    id: 'XLOOKUP', zh: '進階查找', desc: '新一代查找函數，可向任何方向搜尋，支援找不到時的預設值',
    cat: 'advanced', diff: 'i', ver: '365',
    syntax: '=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])',
    args: [
      { n: 'lookup_value', d: '要查找的值', req: true },
      { n: 'lookup_array', d: '查找範圍（單欄或單列）', req: true },
      { n: 'return_array', d: '要回傳的範圍', req: true },
      { n: 'if_not_found', d: '找不到時的替代值', req: false },
      { n: 'match_mode', d: '0=精確，-1=精確或較小，1=精確或較大，2=萬用字元', req: false },
      { n: 'search_mode', d: '1=從頭搜尋，-1=從尾搜尋，2=二分搜尋', req: false }
    ],
    itype: 'none',
    ex: '=XLOOKUP(A2, 員工ID欄, 姓名欄, "查無此人") → 比 VLOOKUP 更強大',
    tips: ['可向左查找（VLOOKUP 做不到）', '可一次回傳多欄', '內建錯誤處理，不需 IFERROR 包住', '取代 VLOOKUP + IFERROR 的最佳方案'],
    rel: ['VLOOKUP', 'INDEX', 'MATCH', 'XMATCH']
  },
  {
    id: 'XMATCH', zh: '進階位置查找', desc: 'MATCH 的升級版，支援更多搜尋模式與萬用字元',
    cat: 'advanced', diff: 'i', ver: '365',
    syntax: '=XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])',
    args: [
      { n: 'lookup_value', d: '要查找的值', req: true },
      { n: 'lookup_array', d: '查找範圍', req: true },
      { n: 'match_mode', d: '0=精確，-1=精確或較小，1=精確或較大，2=萬用字元', req: false },
      { n: 'search_mode', d: '1=從頭，-1=從尾，2/−2=二分搜尋', req: false }
    ],
    itype: 'none',
    ex: '=XMATCH("李*", A:A, 2) → 找第一個姓李的位置（萬用字元）',
    tips: ['可從尾部搜尋（找最後出現的位置）', '配合 INDEX 使用取代 INDEX+MATCH', '支援萬用字元 match_mode=2'],
    rel: ['MATCH', 'XLOOKUP', 'INDEX']
  },
  {
    id: 'FILTER', zh: '動態篩選', desc: '依條件篩選範圍，直接溢出顯示所有符合結果',
    cat: 'advanced', diff: 'i', ver: '365',
    syntax: '=FILTER(array, include, [if_empty])',
    args: [
      { n: 'array', d: '要篩選的範圍', req: true },
      { n: 'include', d: '篩選條件（回傳 TRUE/FALSE 的陣列）', req: true },
      { n: 'if_empty', d: '無符合結果時的顯示值', req: false }
    ],
    itype: 'none',
    ex: '=FILTER(A:C, B:B="業務", "無資料") → 顯示所有業務部門員工',
    tips: ['結果自動溢出（Spill），不需 Ctrl+Shift+Enter', '多條件：=FILTER(A:C,(B:B="業務")*(C:C>50000))', '可搭配 SORT 排序結果'],
    rel: ['SORT', 'UNIQUE', 'XLOOKUP']
  },
  {
    id: 'UNIQUE', zh: '去重複值', desc: '從範圍中取出唯一值（去除重複），自動溢出',
    cat: 'advanced', diff: 'i', ver: '365',
    syntax: '=UNIQUE(array, [by_col], [exactly_once])',
    args: [
      { n: 'array', d: '要去重的範圍', req: true },
      { n: 'by_col', d: 'FALSE=按列去重（預設），TRUE=按欄去重', req: false },
      { n: 'exactly_once', d: 'TRUE=只回傳出現恰好一次的值', req: false }
    ],
    itype: 'none',
    ex: '=UNIQUE(A:A) → 取出A欄所有不重複的值',
    tips: ['結果自動溢出，不需複製公式', '配合 SORT：=SORT(UNIQUE(A:A)) 取得排序的唯一值', '=COUNTA(UNIQUE(A2:A100)) 計算不重複項目數'],
    rel: ['FILTER', 'SORT', 'COUNTIF']
  },
  {
    id: 'SORT', zh: '動態排序', desc: '對範圍進行排序並直接溢出結果',
    cat: 'advanced', diff: 'i', ver: '365',
    syntax: '=SORT(array, [sort_index], [sort_order], [by_col])',
    args: [
      { n: 'array', d: '要排序的範圍', req: true },
      { n: 'sort_index', d: '排序依據的欄/列號（預設第1欄）', req: false },
      { n: 'sort_order', d: '1=升冪（預設），-1=降冪', req: false },
      { n: 'by_col', d: 'FALSE=按列排（預設），TRUE=按欄排', req: false }
    ],
    itype: 'none',
    ex: '=SORT(A1:C10, 2, -1) → 依第2欄降冪排序',
    tips: ['結果自動溢出，原始資料不改動', '=SORT(UNIQUE(A:A)) 取得排序的唯一清單', 'SORTBY 可用其他範圍作為排序依據'],
    rel: ['SORTBY', 'FILTER', 'UNIQUE']
  },
  {
    id: 'SORTBY', zh: '依條件排序', desc: '依一個或多個其他範圍的值進行排序',
    cat: 'advanced', diff: 'i', ver: '365',
    syntax: '=SORTBY(array, by_array1, [sort_order1], [by_array2, sort_order2], ...)',
    args: [
      { n: 'array', d: '要排序的範圍', req: true },
      { n: 'by_array1', d: '排序依據範圍', req: true },
      { n: 'sort_order1', d: '1=升冪，-1=降冪', req: false }
    ],
    itype: 'none',
    ex: '=SORTBY(A:A, B:B, -1) → 依B欄數值降冪排序A欄',
    tips: ['可多欄排序：先按部門，再按薪資', '排序依據範圍不需要在回傳陣列中'],
    rel: ['SORT', 'FILTER', 'UNIQUE']
  },
  {
    id: 'SEQUENCE', zh: '產生序列', desc: '產生連續數字序列（行×欄的陣列）',
    cat: 'advanced', diff: 'i', ver: '365',
    syntax: '=SEQUENCE(rows, [cols], [start], [step])',
    args: [
      { n: 'rows', d: '要產生的列數', req: true },
      { n: 'cols', d: '要產生的欄數（預設1）', req: false },
      { n: 'start', d: '起始值（預設1）', req: false },
      { n: 'step', d: '每步增量（預設1）', req: false }
    ],
    itype: 'none',
    ex: '=SEQUENCE(5) → 1, 2, 3, 4, 5（溢出5列）\n=SEQUENCE(3, 3, 0, 10) → 3×3的矩陣，從0開始每步加10',
    tips: ['可取代 ROW(A1:A10)-1 等舊寫法', '=TEXT(SEQUENCE(12), "00") 產生 01~12 的月份清單', '=DATE(2024, SEQUENCE(12), 1) 一次產生全年月份'],
    rel: ['ROWS', 'ROW', 'FILTER']
  },
  {
    id: 'LET', zh: '命名變數', desc: '在公式中定義命名變數，讓複雜公式更易讀、效能更好',
    cat: 'advanced', diff: 'a', ver: '365',
    syntax: '=LET(name1, value1, [name2, value2], ..., calculation)',
    args: [
      { n: 'name1', d: '第一個變數名稱', req: true },
      { n: 'value1', d: '第一個變數的值或公式', req: true },
      { n: '...calculation', d: '最後一個引數是使用這些變數的計算公式', req: true }
    ],
    itype: 'none',
    ex: '=LET(稅前, A1, 稅率, 0.2, 稅前*(1-稅率)) → 比直接寫 =A1*(1-0.2) 更易讀',
    tips: ['同一個複雜計算在公式中只需寫一次，效能更好', '變數名稱不可含空格', '適合簡化超長公式'],
    rel: ['LAMBDA', 'REDUCE', 'MAP']
  },
  {
    id: 'LAMBDA', zh: '自訂函數', desc: '建立可重複使用的自訂函數（存在名稱管理員中）',
    cat: 'advanced', diff: 'a', ver: '365',
    syntax: '=LAMBDA([parameter1, ...], calculation)',
    args: [
      { n: 'parameter1, ...', d: '函數的輸入參數名稱', req: false },
      { n: 'calculation', d: '使用這些參數的計算公式', req: true }
    ],
    itype: 'none',
    ex: '建立 TAXAMT = LAMBDA(income, rate, income*rate)\n→ 之後可直接用 =TAXAMT(A1, 0.2)',
    tips: ['在名稱管理員定義，之後像內建函數一樣使用', '支援遞迴（函數呼叫自己）', 'MAP/REDUCE/SCAN 等高階函數會用到 LAMBDA'],
    rel: ['LET', 'MAP', 'REDUCE', 'BYROW']
  },
  {
    id: 'MAP', zh: '映射轉換', desc: '對陣列每個元素套用 LAMBDA 函數，回傳轉換後的陣列',
    cat: 'advanced', diff: 'a', ver: '365',
    syntax: '=MAP(array1, [array2, ...], lambda)',
    args: [
      { n: 'array1', d: '要轉換的陣列', req: true },
      { n: 'lambda', d: '套用於每個元素的 LAMBDA 函數', req: true }
    ],
    itype: 'none',
    ex: '=MAP(A1:A10, LAMBDA(x, x*1.1)) → 將A欄所有數值乘以1.1（調薪10%）',
    tips: ['比 BYROW/BYCOL 更靈活', '可傳入多個陣列（對應位置一起處理）', 'Excel 的函數式程式設計入門'],
    rel: ['LAMBDA', 'REDUCE', 'BYROW', 'FILTER']
  },
  {
    id: 'REDUCE', zh: '累積計算', desc: '用 LAMBDA 對陣列每個元素進行累積計算，最終回傳單一結果',
    cat: 'advanced', diff: 'a', ver: '365',
    syntax: '=REDUCE([initial_value], array, lambda)',
    args: [
      { n: 'initial_value', d: '累積的初始值', req: false },
      { n: 'array', d: '要處理的陣列', req: true },
      { n: 'lambda', d: 'LAMBDA(accumulator, current_value, 計算)', req: true }
    ],
    itype: 'none',
    ex: '=REDUCE(0, A1:A5, LAMBDA(acc, x, acc+x)) → 手動實作 SUM',
    tips: ['用於 SUM/MAX/MIN 無法直接表達的複雜累積邏輯', '搭配 LAMBDA 自訂累積方式', '進階用途：自訂加權平均、自訂複利計算'],
    rel: ['LAMBDA', 'MAP', 'SCAN']
  },

];

// ═══════════════════════════════════════════════
//  FORMULA COMBOS  (實戰組合)
// ═══════════════════════════════════════════════
window.COMBOS = [
  {
    id: 'iferror-vlookup',
    title: 'IFERROR + VLOOKUP',
    subtitle: '安全查找：查無結果時顯示自訂訊息',
    cat: 'combo',
    code: `=IFERROR(
  VLOOKUP(A2, 員工資料!A:C, 2, FALSE),
  "查無此人"
)`,
    explanation: 'VLOOKUP 查無結果時回傳 #N/A 錯誤，IFERROR 將其替換為友善訊息。',
    whenToUse: '所有使用 VLOOKUP 的場景，避免讓使用者看到錯誤代碼。',
    upgrade: '改用 XLOOKUP（365），內建錯誤處理參數，不需 IFERROR。',
    diff: 'i', ver: 'all'
  },
  {
    id: 'index-match',
    title: 'INDEX + MATCH',
    subtitle: '萬用查找：可向任何方向，不限欄位順序',
    cat: 'combo',
    code: `=INDEX(
  B:B,
  MATCH(A2, A:A, 0)
)`,
    explanation: 'MATCH 找出 A2 在 A 欄的位置，INDEX 用這個位置從 B 欄取值。可向左查找，VLOOKUP 做不到。',
    whenToUse: '需要向左查找、回傳欄位靠左、或不想因插入欄而破壞欄號時。',
    upgrade: '在 Excel 365 直接用 XLOOKUP，語法更簡潔。',
    diff: 'i', ver: 'all'
  },
  {
    id: 'iferror-index-match',
    title: 'IFERROR + INDEX + MATCH',
    subtitle: '最強容錯查找組合（Excel 2016 以下的 XLOOKUP 替代方案）',
    cat: 'combo',
    code: `=IFERROR(
  INDEX(B:B, MATCH(A2, A:A, 0)),
  "查無資料"
)`,
    explanation: '結合 INDEX+MATCH 的靈活查找與 IFERROR 的錯誤處理，是 XLOOKUP 在舊版 Excel 的最佳替代方案。',
    whenToUse: '需要靈活查找且必須相容 Excel 2016 以下版本。',
    upgrade: 'Excel 365 請直接使用 XLOOKUP。',
    diff: 'i', ver: 'all'
  },
  {
    id: 'if-and',
    title: 'IF + AND（多條件）',
    subtitle: '同時滿足多個條件才執行',
    cat: 'combo',
    code: `=IF(
  AND(B2="業務部", C2>=80000),
  "達標獎金",
  "未達標"
)`,
    explanation: '用 AND 連接多個條件，所有條件都必須為 TRUE，IF 才回傳第一個值。',
    whenToUse: '需要同時滿足多個條件才觸發某個結果時。',
    upgrade: 'Excel 365 的 IFS 或 SWITCH 對更多條件更清晰。',
    diff: 'b', ver: 'all'
  },
  {
    id: 'if-or',
    title: 'IF + OR（任一條件）',
    subtitle: '符合任一條件即執行',
    cat: 'combo',
    code: `=IF(
  OR(D2="主管", D2="協理", D2="副總"),
  "管理職",
  "一般職"
)`,
    explanation: '用 OR 連接多個可能的值，只要符合其中一個，IF 就回傳第一個值。',
    whenToUse: '多個不同值都對應同一個結果時，比多層 IF 更清晰。',
    diff: 'b', ver: 'all'
  },
  {
    id: 'text-today',
    title: 'TEXT + TODAY',
    subtitle: '把今天的日期格式化成中文或指定樣式',
    cat: 'combo',
    code: `=TEXT(TODAY(), "yyyy年mm月dd日") & " (" & TEXT(TODAY(),"aaaa") & ")"`,
    explanation: 'TEXT 將日期數值轉為指定格式的文字字串，aaaa 格式代碼在繁體中文 Excel 輸出中文星期。',
    whenToUse: '在報告標題、列印版面需要顯示「本日日期+星期」的中文格式時。',
    diff: 'i', ver: 'all'
  },
  {
    id: 'textjoin-if',
    title: 'TEXTJOIN + IF（條件串接）',
    subtitle: '把符合條件的多個值合併為一個字串',
    cat: 'combo',
    code: `=TEXTJOIN(
  "、",
  TRUE,
  IF(B2:B100="業務部", A2:A100, "")
)`,
    explanation: '用 IF 篩選出符合條件的名字（不符合回傳空字串），TEXTJOIN 將它們用頓號連接起來，並自動忽略空字串。',
    whenToUse: '需要把同類別的值合併成一格顯示，例如「同部門員工」、「相同標籤項目」。',
    diff: 'a', ver: '2016'
  },
  {
    id: 'sumifs-date',
    title: 'SUMIFS + 日期條件',
    subtitle: '計算特定時間區間的加總',
    cat: 'combo',
    code: `=SUMIFS(
  C:C,
  A:A, ">=" & DATE(2024, 1, 1),
  A:A, "<=" & DATE(2024, 3, 31)
)`,
    explanation: '使用 ">=" & DATE(...) 的語法設定日期條件的上下限，計算特定期間（如第一季）的業績加總。',
    whenToUse: '月報、季報、年報等需要依日期範圍加總的場景。',
    diff: 'i', ver: 'all'
  },
  {
    id: 'filter-sort',
    title: 'FILTER + SORT（365）',
    subtitle: '篩選並自動排序，結果動態更新',
    cat: 'combo',
    code: `=SORT(
  FILTER(
    A2:C100,
    (B2:B100="業務部") * (C2:C100>=60000)
  ),
  3, -1
)`,
    explanation: 'FILTER 先篩出業務部且薪資≥6萬的員工，SORT 再依第3欄（薪資）由高到低排序。兩個條件相乘（*）等同 AND。',
    whenToUse: '取代需要手動更新的 Excel 篩選，讓清單永遠保持動態最新。',
    diff: 'i', ver: '365'
  },
  {
    id: 'xlookup-two-way',
    title: 'XLOOKUP 雙向查找（365）',
    subtitle: '同時查找列與欄的交叉點',
    cat: 'combo',
    code: `=XLOOKUP(
  G1,
  A1:A10,
  XLOOKUP(H1, B1:F1, B2:F10)
)`,
    explanation: '外層 XLOOKUP 找到對應的列，內層 XLOOKUP 找到對應的欄（回傳整列），兩者結合取得矩陣中的交叉值。取代需要 INDEX+MATCH+MATCH 的複雜寫法。',
    whenToUse: '查找矩陣交叉點，例如依員工+月份查找出勤數、依產品+地區查找銷售額。',
    diff: 'a', ver: '365'
  }
];
