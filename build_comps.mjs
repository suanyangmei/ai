import fs from 'node:fs/promises';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const outputDir = 'outputs';
await fs.mkdir(outputDir, { recursive: true });

const rows = [
  { name:'中控技术', code:'688777.SH', group:'控制与自动化', revenue:9138514049.21, revenueYoy:0.060163412204, netProfit:1116986722.13, netYoy:0.013816927128, rdExpense:978005205.11, rdTotal:978005205.11, rdRate:0.1070, relatedRevenue:5642665000, relatedLabel:'工业自动化及智能制造解决方案', aiRevenue:null, aiLabel:'AI收入未单列', products:'DCS/UCS、工业软件、仪器仪表、运维服务', customers:'石化、化工、电力、冶金、建材', model:'控制系统/产品 + 项目交付 + 运维', caution:'5.643亿元为更宽的自动化及智能制造解决方案收入，不等于工业 AI 收入', pages:'年报 pp.9-10、47、59-60', pdf:'https://www.sse.com.cn/disclosure/listedinfo/announcement/c/new/2025-04-01/688777_20250401_F6Z3.pdf', codeApi:'https://emweb.securities.eastmoney.com/PC_HSF10/NewFinanceAnalysis/lrbAjaxNew?companyType=4&reportDateType=0&reportType=1&dates=2024-12-31&code=SH688777'},
  { name:'汇川技术', code:'300124.SZ', group:'控制与自动化', revenue:37040952138.72, revenueYoy:0.217654272801, netProfit:4285493340.70, netYoy:-0.096242757980, rdExpense:3147080701.82, rdTotal:3147080701.82, rdRate:0.0850, relatedRevenue:15242986575.24, relatedLabel:'通用自动化类', aiRevenue:null, aiLabel:'AI收入未单列', products:'变频、伺服、PLC、通用自动化、新能源汽车电驱', customers:'装备制造、锂电、光伏、汽车、物流', model:'标准产品 + 行业解决方案', caution:'通用自动化类含控制与驱动产品，不等于 AI 业务收入', pages:'年报 pp.9-10、28、34-35', pdf:'https://pdf.dfcfw.com/pdf/H2_AN202504281664018483_1.pdf', codeApi:'https://emweb.securities.eastmoney.com/PC_HSF10/NewFinanceAnalysis/lrbAjaxNew?companyType=4&reportDateType=0&reportType=1&dates=2024-12-31&code=SZ300124'},
  { name:'宝信软件', code:'600845.SH', group:'工业软件/平台', revenue:13644424669.39, revenueYoy:0.056426408471, netProfit:2265482898.16, netYoy:-0.112841351053, rdExpense:1461458198.33, rdTotal:1461458198.33, rdRate:0.1071, relatedRevenue:9910630279.83, relatedLabel:'软件开发及工程服务', aiRevenue:null, aiLabel:'AI收入未单列', products:'工业软件、自动化、软件工程、服务外包、系统集成、IDC', customers:'钢铁、制造、政企和云客户', model:'软件/系统项目 + 服务外包 + 云服务', caution:'软件开发及工程服务涵盖多个行业与项目，不能视为工业 AI 收入', pages:'年报 pp.6、14-16', pdf:'https://www.sse.com.cn/disclosure/listedinfo/announcement/c/new/2025-04-01/600845_20250401_EA1P.pdf', codeApi:'https://emweb.securities.eastmoney.com/PC_HSF10/NewFinanceAnalysis/lrbAjaxNew?companyType=4&reportDateType=0&reportType=1&dates=2024-12-31&code=SH600845'},
  { name:'能科科技', code:'603859.SH', group:'工业软件/服务', revenue:1509822499.58, revenueYoy:0.074745578971, netProfit:191734343.34, netYoy:-0.151709480708, rdExpense:194428670.42, rdTotal:455783153.04, rdRate:0.3019, relatedRevenue:273002618.85, relatedLabel:'AI产品与服务', aiRevenue:273002618.85, aiLabel:'年报单列AI产品与服务', products:'AI产品与服务、云产品与服务、软件系统与服务、工业工程及工业电气', customers:'重工装备、汽车、高科技电子、通用机械、能源动力', model:'软件/产品 + 数字化工程项目 + 云服务', caution:'AI产品与服务包含多类AI应用，不全等同工业 AI；研发投入含资本化开发支出', pages:'年报 pp.6-7、15-16、19', pdf:'https://www.sse.com.cn/disclosure/listedinfo/announcement/c/new/2025-04-25/603859_20250425_7RLG.pdf', codeApi:'https://emweb.securities.eastmoney.com/PC_HSF10/NewFinanceAnalysis/lrbAjaxNew?companyType=4&reportDateType=0&reportType=1&dates=2024-12-31&code=SH603859'},
  { name:'赛意信息', code:'300687.SZ', group:'制造数字化服务', revenue:2395451665.47, revenueYoy:0.062744924787, netProfit:139378011.21, netYoy:-0.452137920320, rdExpense:262773844.42, rdTotal:364985576.88, rdRate:0.1524, relatedRevenue:990252826.45, relatedLabel:'智能制造及工业互联网', aiRevenue:null, aiLabel:'AI收入未单列', products:'泛 ERP、智能制造及工业互联网、软件维护、代理分销', customers:'电子、家电、汽车、装备、消费品制造', model:'软件实施/开发 + 软件产品 + 技术服务', caution:'该业务线包含数字化和工业互联网项目，年报未将其中AI收入单列；研发投入与研发费用口径不同', pages:'年报 pp.8-9、31、34-36', pdf:'https://pdf.dfcfw.com/pdf/H2_AN202504241661820324_1.pdf', codeApi:'https://emweb.securities.eastmoney.com/PC_HSF10/NewFinanceAnalysis/lrbAjaxNew?companyType=4&reportDateType=0&reportType=1&dates=2024-12-31&code=SZ300687'},
  { name:'海康威视', code:'002415.SZ', group:'感知与机器视觉', revenue:92495525118.30, revenueYoy:0.035306759909, netProfit:11977327023.54, netYoy:-0.151009398042, rdExpense:11864013392.27, rdTotal:11864013392.27, rdRate:0.1283, relatedRevenue:null, relatedLabel:'未披露制造/工业AI单独收入', aiRevenue:null, aiLabel:'AI收入未按工业场景单列', products:'主业产品及服务、机器人、智能家居、热成像、汽车电子、存储等创新业务', customers:'政企、制造、交通、能源、物流、消费市场', model:'感知硬件 + 软件平台 + 行业方案', caution:'创新业务及机器人业务覆盖多种场景，不可等同制造业AI收入', pages:'年报 pp.9、89-90、94-95', pdf:'https://pdf.dfcfw.com/pdf/H2_AN202504181658527316_1.pdf', codeApi:'https://emweb.securities.eastmoney.com/PC_HSF10/NewFinanceAnalysis/lrbAjaxNew?companyType=4&reportDateType=0&reportType=1&dates=2024-12-31&code=SZ002415'},
];

const wb = Workbook.create();
const summary = wb.worksheets.add('对比摘要');
const finance = wb.worksheets.add('财务数据');
const business = wb.worksheets.add('业务定位与收入结构');
for (const s of [summary, finance, business]) { s.showGridLines = false; s.getRange('A1:Z200').format.font = { name: 'Arial', size: 10, color: '#1F2937' }; }

summary.getRange('A1:M1').merge();
summary.getRange('A1').values = [['工业 AI 制造业上市公司对比（2024 年报口径）']];
summary.getRange('A2:M2').merge();
summary.getRange('A2').values = [['金额单位：人民币亿元；研发费用、研发投入合计及年报披露研发强度分列，收入仅按已披露的相关业务分类映射。']];
summary.getRange('A4:M4').values = [['公司','代码','可比组','营业收入（亿元）','收入同比','归母净利润（亿元）','净利润同比','研发费用（亿元）','研发投入合计（亿元）','年报研发强度','研发费用/收入','单列AI产品收入（亿元）','年报页码']];
summary.getRange('A5:C10').values = rows.map(r => [r.name, r.code, r.group]);
summary.getRange('D5:K10').formulas = rows.map((_, i) => [`='财务数据'!D${i+5}/100000000`, `='财务数据'!E${i+5}`, `='财务数据'!F${i+5}/100000000`, `='财务数据'!G${i+5}`, `='财务数据'!H${i+5}/100000000`, `='财务数据'!I${i+5}/100000000`, `='财务数据'!J${i+5}`, `='财务数据'!K${i+5}`]);
summary.getRange('L5:L10').values = rows.map(r => [r.aiRevenue === null ? '未单列' : r.aiRevenue / 1e8]);
summary.getRange('M5:M10').values = rows.map(r => [r.pages]);
summary.getRange('A12:M12').merge();
summary.getRange('A12').values = [['能科科技年报单列AI产品与服务收入2.73亿元；该口径包含多类AI应用，不全等同工业AI。其余公司未单列工业AI收入。宽口径相关业务见“业务定位与收入结构”页，不得当作工业AI收入。']];

finance.getRange('A1:M1').merge();
finance.getRange('A1').values = [['财务数据与来源（2024-12-31 年报）']];
finance.getRange('A2:M2').merge();
finance.getRange('A2').values = [['财务金额及研发数据优先按年报合并口径录入；F10接口作为交叉校验。来源页码列按PDF印刷页标示。']];
finance.getRange('A4:M4').values = [['公司','代码','报告期','营业收入（元）','收入同比','归母净利润（元）','净利润同比','研发费用（元）','研发投入合计（元）','研发投入占收入（年报）','研发费用/收入（公式）','年报页码','年报PDF链接']];
finance.getRange('A5:C10').values = rows.map(r => [r.name, r.code, '2024-12-31']);
finance.getRange('D5:J10').values = rows.map(r => [r.revenue, r.revenueYoy, r.netProfit, r.netYoy, r.rdExpense, r.rdTotal, r.rdRate]);
finance.getRange('K5:K10').formulas = rows.map((_, i) => [`=H${i+5}/D${i+5}`]);
finance.getRange('L5:M10').values = rows.map(r => [r.pages, r.pdf]);
finance.getRange('A12:M12').merge();
finance.getRange('A12').values = [['研发投入合计按年报“研发投入”表，研发费用按利润表费用项。能科、赛意存在资本化研发支出，两列不可互换；公司披露的研发强度按年报口径列示。']];

business.getRange('A1:I1').merge();
business.getRange('A1').values = [['主营业务、收入结构及工业 AI 收入口径（2024 年报）']];
business.getRange('A2:I2').merge();
business.getRange('A2').values = [['金额单位：人民币亿元；业务收入仅为年报披露的相关产品/业务线，不代表工业 AI 收入。直接AI收入未单列时保持为空。']];
business.getRange('A3:I3').values = [['公司','可比组','年报收入结构（主要分类）','相关业务线','相关业务收入（亿元）','收入占比','单列AI产品收入（亿元）','商业模式/下游','边界及来源页码']];
const businessRows = [
  ['中控技术','控制与自动化','自动化/智能制造解决方案、仪器仪表、工业软件、运维、S2B','工业自动化及智能制造解决方案',56.42665, null, null,'控制系统/产品 + 项目交付 + 运维；流程工业','更宽自动化业务，不是AI收入；年报 pp.59-60'],
  ['汇川技术','控制与自动化','通用自动化、智慧电梯、新能源汽车及轨道交通、其他','通用自动化类',152.4298657524, null, null,'标准产品 + 行业方案；装备/新能源制造','包含非AI自动化产品；年报 p.28'],
  ['宝信软件','工业软件/平台','软件开发及工程服务、服务外包、系统集成','软件开发及工程服务',99.1063027983, null, null,'软件项目/集成 + 服务外包 + 云服务；钢铁/制造/政企','软件工程业务跨行业，不能视为AI收入；年报 p.14'],
  ['能科科技','工业软件/服务','AI产品与服务、云产品与服务、软件系统与服务、工业工程及工业电气','AI产品与服务',2.7300261885, null, 2.7300261885,'软件/产品 + 数字化工程项目 + 云服务；重工/汽车/高科技电子/通用机械/能源动力','AI产品线不全等同工业AI；年报 pp.15-16'],
  ['赛意信息','制造数字化服务','泛ERP、智能制造及工业互联网、软件维护、代理分销、其他','智能制造及工业互联网',9.9025282645, null, null,'软件实施/开发 + 产品 + 技术服务；电子/家电/汽车等制造','该线包含数字化项目，未单列其中AI收入；年报 p.31'],
  ['海康威视','感知与机器视觉','主业产品与服务、主业建造工程、机器人/智能家居/热成像/汽车电子/存储等创新业务','未单列制造或工业AI业务线', null, null, null,'感知硬件 + 软件平台 + 行业方案；政企/制造/交通/能源/物流','机器人及创新业务不等于制造AI；年报 pp.89-90'],
];
business.getRange('A4:I9').values = businessRows.map((r, i) => [r[0], r[1], r[2], r[3], r[4], null, r[6], r[7], r[8]]);
business.getRange('F4:F9').formulas = rows.map((_, i) => [`=IF(E${i+4}="","",E${i+4}*100000000/'财务数据'!D${i+5})`]);
business.getRange('E4:E9').format.numberFormat = '0.00';
business.getRange('F4:F9').format.numberFormat = '0.00%';
business.getRange('G4:G9').format.numberFormat = '0.00';

for (const sheet of [summary, finance, business]) {
  sheet.getRange('A1:Z1').format = { font: { name: 'Arial', size: 14, bold: true, color: '#123B5D' } };
  sheet.getRange('A2:Z2').format = { font: { name: 'Arial', size: 9, italic: true, color: '#5B6470' }, wrapText: true };
  sheet.getRange('A1:Z2').format.rowHeight = 22;
}
for (const range of [summary.getRange('A4:M4'), finance.getRange('A4:M4'), business.getRange('A3:I3')]) {
  range.format = { fill: '#123B5D', font: { name: 'Arial', size: 10, bold: true, color: '#FFFFFF' }, wrapText: true, horizontalAlignment: 'center', verticalAlignment: 'center' };
  range.format.rowHeight = 30;
  range.format.borders = { preset: 'outside', style: 'thin', color: '#123B5D' };
}
summary.getRange('A5:M10').format.borders = { preset: 'inside', style: 'thin', color: '#D9E2EC' };
finance.getRange('A5:M10').format.borders = { preset: 'inside', style: 'thin', color: '#D9E2EC' };
business.getRange('A4:I9').format.borders = { preset: 'inside', style: 'thin', color: '#D9E2EC' };
summary.getRange('G5:J10').format.wrapText = true;
summary.getRange('A5:M10').format.rowHeight = 38;
finance.getRange('L5:M10').format.wrapText = true;
finance.getRange('M5:M10').format.font = { name: 'Arial', size: 8, color: '#2563EB' };
finance.getRange('A5:M10').format.rowHeight = 34;
business.getRange('C4:I9').format.wrapText = true;
business.getRange('A4:I9').format.rowHeight = 42;
summary.getRange('D5:D10').format.numberFormat = '0.00';
summary.getRange('E5:E10').format.numberFormat = '0.00%';
summary.getRange('F5:F10').format.numberFormat = '0.00';
summary.getRange('G5:G10').format.numberFormat = '0.00%';
summary.getRange('H5:I10').format.numberFormat = '0.00';
summary.getRange('J5:K10').format.numberFormat = '0.00%';
summary.getRange('L5:L10').format.numberFormat = '0.00';
finance.getRange('D5:D10').format.numberFormat = '#,##0.00';
finance.getRange('E5:E10').format.numberFormat = '0.00%';
finance.getRange('F5:F10').format.numberFormat = '#,##0.00';
finance.getRange('G5:G10').format.numberFormat = '0.00%';
finance.getRange('H5:I10').format.numberFormat = '#,##0.00';
finance.getRange('J5:K10').format.numberFormat = '0.00%';
summary.getRange('A5:M10').format.verticalAlignment = 'center';
finance.getRange('A5:M10').format.verticalAlignment = 'center';
business.getRange('A4:I9').format.verticalAlignment = 'center';
for (const s of [summary, finance, business]) { s.getUsedRange().format.autofitColumns(); s.getUsedRange().format.autofitRows(); }
summary.getRange('A:A').format.columnWidth = 14; summary.getRange('B:B').format.columnWidth = 12; summary.getRange('C:C').format.columnWidth = 16; summary.getRange('D:G').format.columnWidth = 16; summary.getRange('H:I').format.columnWidth = 19; summary.getRange('J:K').format.columnWidth = 18; summary.getRange('L:L').format.columnWidth = 23; summary.getRange('M:M').format.columnWidth = 38;
finance.getRange('A:A').format.columnWidth = 14; finance.getRange('B:B').format.columnWidth = 12; finance.getRange('C:D').format.columnWidth = 18; finance.getRange('E:K').format.columnWidth = 18; finance.getRange('L:L').format.columnWidth = 30; finance.getRange('M:M').format.columnWidth = 58;
business.getRange('A:A').format.columnWidth = 14; business.getRange('B:B').format.columnWidth = 16; business.getRange('C:C').format.columnWidth = 30; business.getRange('D:D').format.columnWidth = 38; business.getRange('E:E').format.columnWidth = 22; business.getRange('F:F').format.columnWidth = 20; business.getRange('G:G').format.columnWidth = 20; business.getRange('H:H').format.columnWidth = 26; business.getRange('I:I').format.columnWidth = 40;
summary.freezePanes.freezeRows(4); finance.freezePanes.freezeRows(4); business.freezePanes.freezeRows(3);

wb.recalculate();
const inspect = await wb.inspect({ kind: 'table', sheetId: '对比摘要', range: 'A1:M12', include: 'values,formulas', tableMaxRows: 15, tableMaxCols: 13, maxChars: 12000 });
await fs.writeFile(`${outputDir}/对比摘要.inspect.json`, inspect.ndjson, 'utf8');
const financeInspect = await wb.inspect({ kind: 'table', sheetId: '财务数据', range: 'A4:M10', include: 'values,formulas', tableMaxRows: 10, tableMaxCols: 13, maxChars: 12000 });
await fs.writeFile(`${outputDir}/财务数据.inspect.json`, financeInspect.ndjson, 'utf8');
const errors = await wb.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!', options: { useRegex: true, maxResults: 100 }, summary: 'formula error scan' });
await fs.writeFile(`${outputDir}/公式错误检查.inspect.json`, errors.ndjson, 'utf8');
const png = await wb.render({ sheetName: '对比摘要', range: 'A1:M12', scale: 1, format: 'png' });
await fs.writeFile(`${outputDir}/对比摘要.png`, new Uint8Array(await png.arrayBuffer()));
const financePng = await wb.render({ sheetName: '财务数据', range: 'A1:M12', scale: 1, format: 'png' });
await fs.writeFile(`${outputDir}/财务数据.png`, new Uint8Array(await financePng.arrayBuffer()));
const businessPng = await wb.render({ sheetName: '业务定位与收入结构', range: 'A1:I10', scale: 1, format: 'png' });
await fs.writeFile(`${outputDir}/业务定位与收入结构.png`, new Uint8Array(await businessPng.arrayBuffer()));
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(`${outputDir}/公司对比表.xlsx`);
