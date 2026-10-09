// Rebuilds the public synthetic example. Development requires @oai/artifact-tool.
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

// A deliberately cross-linked, but acyclic, synthetic operating workbook.
const repoDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const dir = path.join(repoDir, 'examples');
const qaArg = process.argv.indexOf('--qa-dir');
const qaDir = qaArg >= 0 ? path.resolve(process.argv[qaArg + 1]) : path.join(os.tmpdir(), 'excel-formula-atlas-example-qa');
await fs.mkdir(dir, { recursive: true });
await fs.mkdir(qaDir, { recursive: true });
await fs.mkdir(path.join(repoDir, 'docs', 'downloads'), { recursive: true });
console.log('Building the eight-sheet synthetic operations model.');
const wb = Workbook.create();
const sheetNames = ['Summary','Assumptions','Sales','Costing','Purchasing','Inventory','Orders','Catalog'];
const sheets = Object.fromEntries(sheetNames.map(n=>[n,wb.worksheets.add(n)]));
const expected = [], cells = [], regions = [];
const numFormat = '#,##0;(#,##0);"-"';
const sourceColor = '#2155B5', linkColor = '#167044';
const products = [
  {sku:'A100',name:'Sensor mount',price:85000,mat:'R01',use:2,labor:12000,safety:20,opening:35},
  {sku:'A200',name:'Mobile stand',price:140000,mat:'R01',use:3,labor:18000,safety:12,opening:18},
  {sku:'B100',name:'Control box',price:220000,mat:'R02',use:1,labor:26000,safety:10,opening:14},
  {sku:'B200',name:'Weatherproof enclosure',price:310000,mat:'R03',use:2,labor:34000,safety:8,opening:10},
];
const materials = [
  {code:'R01',name:'Aluminum frame',unit:18000,lead:7,supplier:'North Metals',opening:45,safety:20},
  {code:'R02',name:'Control module',unit:76000,lead:14,supplier:'Bright Electronics',opening:7,safety:4},
  {code:'R03',name:'Sealed case',unit:68000,lead:10,supplier:'Ocean Parts',opening:12,safety:6},
];
const qty = [12,9,8,7,18,14,11,9,15,10,12,8,21,16,13,11];
const customers = ['Future Machines','West Industries','Steady Equipment','Grove Automation'];
const orders = qty.map((q,i)=>({id:`O-${String(i+1).padStart(3,'0')}`,day:i+2,customer:customers[Math.floor(i/4)],sku:products[i%4].sku,qty:q,discount:[.03,.05,.02,.04][Math.floor(i/4)],returns:[3,8,14].includes(i)?1:0}));
const settings = {collect:.8,oldCollect:8500000,pay:.65,fixed:4200000,capex:2800000,beginCash:32000000,minCash:12000000};

function V(s,a,v,label='') {
  sheets[s].getRange(a).values = [[v]];
  if (typeof v==='number') { sheets[s].getRange(a).setNumberFormat(numFormat); sheets[s].getRange(a).format.font.color=sourceColor; }
  cells.push({sheet:s,address:a,label,value:v,formula:''});
}
function F(s,a,formula,value,label='') {
  sheets[s].getRange(a).formulas=[[formula]];
  sheets[s].getRange(a).setNumberFormat(numFormat);
  sheets[s].getRange(a).format.font.color=s==='Summary'?'#132B3B':formula.includes('!')?linkColor:'#172A39';
  expected.push({sheet:s,address:a,expected:value,formula,label});
  cells.push({sheet:s,address:a,label,value,formula});
}
function block(s,r,values) { sheets[s].getRange(r).values=values; }
function header(s,r,values) {
  block(s,r,[values]);
  sheets[s].getRange(r).format={fill:'#213D50',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',verticalAlignment:'center',rowHeight:30};
}
function title(s,a,text,subA,sub) {
  V(s,a,text); sheets[s].getRange(a).format.font={name:'Arial',size:16,bold:true,color:'#132B3B'};
  if(sub){V(s,subA,sub);sheets[s].getRange(subA).format.font={name:'Arial',size:10,color:'#536B7A',italic:true};}
}
function region(sheet,range,title,description,role){regions.push({sheet,range,title,description,role});}
function section(s,a,text){V(s,a,text);sheets[s].getRange(a).format.font={name:'Arial',size:11,bold:true,color:'#164D68'};}
function labelcol(s,start,labels){let [_,col,row]=start.match(/^([A-Z]+)(\d+)$/);labels.forEach((l,i)=>V(s,`${col}${+row+i}`,l));}
for(const s of sheetNames){
  const sh=sheets[s]; sh.showGridLines=false;
  sh.getRange('A1:U40').format={font:{name:'Arial',size:10,color:'#172A39'},rowHeight:24,verticalAlignment:'center'};
  sh.getRange('A1:U40').format.columnWidth=12;
  sh.getRange('A1:A40').format.columnWidth=3;
}
sheets['Summary'].tabColor='#183D52'; sheets['Assumptions'].tabColor='#618A9D';

// Central source controls: no downstream result feeds an assumption.
title('Assumptions','D3','October 2026 operating assumptions','D4','All data is synthetic. Monetary amounts are in KRW.');
header('Assumptions','D6:F6',['Assumption','Unit','Input']);
const settingRows=[['Model month','Month','2026-10'],['Current sales collection rate','%',settings.collect],['Prior receivables collected','KRW',settings.oldCollect],['Current purchase payment rate','%',settings.pay],['Monthly fixed operating costs','KRW',settings.fixed],['Capital expenditure paid','KRW',settings.capex],['Opening cash','KRW',settings.beginCash],['Minimum cash balance','KRW',settings.minCash]];
settingRows.forEach((r,i)=>r.forEach((x,j)=>V('Assumptions',`${['D','E','F'][j]}${i+7}`,x,r[0])));
sheets['Assumptions'].getRange('D6:D14').format.columnWidth=26;
sheets['Assumptions'].getRange('F7:F14').format.columnWidth=19;
sheets['Assumptions'].getRange('F8').setNumberFormat('0.0%');sheets['Assumptions'].getRange('F10').setNumberFormat('0.0%');
sheets['Assumptions'].getRange('F7:F14').format.fill='#FFF3C6';
section('Assumptions','D20','Opening material inventory'); header('Assumptions','D21:F21',['Material code','Opening units','Safety units']);
materials.forEach((m,i)=>{V('Assumptions',`D${22+i}`,m.code);V('Assumptions',`E${22+i}`,m.opening,'Opening material units');V('Assumptions',`F${22+i}`,m.safety,'Material safety stock');});
section('Assumptions','M7','Model conventions');
['Revenue excludes returned units.','Discounts are entered on each order.','Unsold production remains in inventory.','Cost of goods sold follows units sold.','All current production labor is paid this month.','65% of material purchases is paid this month.','VAT, income tax and depreciation are excluded.'].forEach((t,i)=>V('Assumptions',`M${9+i}`,t));
region('Assumptions','D6:F14','Cash and operating assumptions','Enter collection and payment rates and opening cash here.','input');
region('Assumptions','D21:F24','Opening material stock','Purchasing deducts this stock from the material requirement.','input');

// Reference master, deliberately placed upper-right and lower-left.
title('Catalog','H3','Product and material catalog','H4','Synthetic prices, material usage, labor costs and inventory targets.');
header('Catalog','H8:O8',['Product code','Product name','List price','Material code','Units per product','Labor per unit','Safety units','Opening units']);
products.forEach((p,i)=>[p.sku,p.name,p.price,p.mat,p.use,p.labor,p.safety,p.opening].forEach((v,j)=>V('Catalog',`${['H','I','J','K','L','M','N','O'][j]}${9+i}`,v,['Product code','Product name','List price','Material code','Material units per product','Labor per unit','Safety stock','Opening stock'][j])));
sheets['Catalog'].getRange('I8:I12').format.columnWidth=19;sheets['Catalog'].getRange('J8:J12').format.columnWidth=16;sheets['Catalog'].getRange('M8:O12').format.columnWidth=17;
sheets['Catalog'].getRange('K9:K12').format.horizontalAlignment='center';
section('Catalog','C26','Material prices and suppliers');header('Catalog','C28:G28',['Material code','Material name','Unit price','Lead days','Supplier']);
materials.forEach((m,i)=>[m.code,m.name,m.unit,m.lead,m.supplier].forEach((v,j)=>V('Catalog',`${['C','D','E','F','G'][j]}${29+i}`,v)));
sheets['Catalog'].getRange('D28:D31').format.columnWidth=22;sheets['Catalog'].getRange('E28:E31').format.columnWidth=16;sheets['Catalog'].getRange('G28:G31').format.columnWidth=18;
sheets['Catalog'].getRange('F29:F31').format.horizontalAlignment='center';
region('Catalog','H8:O12','Product specifications','List prices, material usage, labor costs and finished-goods inventory targets.','input');
region('Catalog','C28:G31','Material specifications','Costing and Purchasing use the same material prices.','input');

// Detailed orders with direct calculation and cross-sheet price retrieval.
title('Orders','B3','October customer orders','B4','Sixteen synthetic orders from four fictional customers, including returns and discounts.');
header('Orders','B7:K7',['Order ID','Order date','Customer','Product code','Ordered units','Discount','Returned units','Notes','Net units sold','Order revenue']);
orders.forEach((o,i)=>{
  const r=i+8,p=products.find(p=>p.sku===o.sku);o.net=o.qty-o.returns;o.revenue=o.net*p.price*(1-o.discount);
  [o.id,new Date(Date.UTC(2026,9,o.day,12)),o.customer,o.sku,o.qty,o.discount,o.returns,o.returns?'Return deducted':'Standard order'].forEach((v,j)=>V('Orders',`${['B','C','D','E','F','G','H','I'][j]}${r}`,v));
  F('Orders',`J${r}`,`=F${r}-H${r}`,o.net,'Net units per order');
  F('Orders',`K${r}`,`=J${r}*SUMIFS('Catalog'!$J$9:$J$12,'Catalog'!$H$9:$H$12,E${r})*(1-G${r})`,o.revenue,'Revenue after discounts and returns');
});
sheets['Orders'].getRange('C8:C23').setNumberFormat('yyyy-mm-dd');sheets['Orders'].getRange('G8:G23').setNumberFormat('0.0%');
sheets['Orders'].getRange('C8:C23').format.horizontalAlignment='center';
sheets['Orders'].getRange('C7:D23').format.columnWidth=16;sheets['Orders'].getRange('I7:I23').format.columnWidth=18;sheets['Orders'].getRange('J7:K23').format.columnWidth=17;
sheets['Orders'].freezePanes.freezeRows(7);
region('Orders','B7:K23','Sixteen customer orders','Deduct returned units and apply each order’s discount to the list price.','source');

// Sales combines source orders, price master and a cash-collection assumption.
title('Sales','E3','Sales and collections by product','E4','Aggregate net units and discounted order revenue by product code.');
header('Sales','E8:L8',['Product code','Net units sold','List price','Monthly revenue','Product name','Cash collected','Receivables','Average selling price']);
products.forEach((p,i)=>{
  const r=9+i,pr=9+i;p.salesQty=orders.filter(o=>o.sku===p.sku).reduce((a,o)=>a+o.net,0);p.sales=orders.filter(o=>o.sku===p.sku).reduce((a,o)=>a+o.revenue,0);p.cash=p.sales*settings.collect;p.receivable=p.sales-p.cash;
  V('Sales',`E${r}`,p.sku);V('Sales',`I${r}`,p.name);
  F('Sales',`F${r}`,`=SUMIFS('Orders'!$J$8:$J$23,'Orders'!$E$8:$E$23,E${r})`,p.salesQty,'Net units by product');
  F('Sales',`G${r}`,`='Catalog'!J${pr}`,p.price,'List price');
  F('Sales',`H${r}`,`=SUMIFS('Orders'!$K$8:$K$23,'Orders'!$E$8:$E$23,E${r})`,p.sales,'Monthly product revenue');
  F('Sales',`J${r}`,`=H${r}*'Assumptions'!$F$8`,p.cash,'Current cash collections');
  F('Sales',`K${r}`,`=H${r}-J${r}`,p.receivable,'Closing receivables');
  F('Sales',`L${r}`,`=H${r}/F${r}`,p.sales/p.salesQty,'Average selling price after discounts');
});
sheets['Sales'].getRange('F8:H12').format.columnWidth=16;sheets['Sales'].getRange('I8:I12').format.columnWidth=19;sheets['Sales'].getRange('J8:L12').format.columnWidth=18;
sheets['Sales'].getRange('I9:I12').format.horizontalAlignment='center';
const totalSales=products.reduce((a,p)=>a+p.sales,0),totalQty=products.reduce((a,p)=>a+p.salesQty,0),cashCollect=totalSales*settings.collect;
section('Sales','P23','Monthly sales and collections');labelcol('Sales','P25',['Net units sold','Monthly revenue','Cash collected','Closing receivables','Average selling price']);
[['Q25','=SUM(F9:F12)',totalQty],['Q26','=SUM(H9:H12)',totalSales],['Q27','=SUM(J9:J12)',cashCollect],['Q28','=SUM(K9:K12)',totalSales-cashCollect],['Q29','=Q26/Q25',totalSales/totalQty]].forEach(([a,f,v])=>F('Sales',a,f,v,a==='Q27'?'Total sales cash collections':''));
sheets['Sales'].getRange('P25:P29').format.columnWidth=21;sheets['Sales'].getRange('Q25:Q29').format.columnWidth=19;
region('Sales','E8:L12','Sales and collections by product','Aggregate order revenue and split it into cash and receivables using the collection rate.','calculation');
region('Sales','P25:Q29','Total revenue','Current collections feed the cash calculation in Summary.','summary');

// Useful costing steps, then quantity times unit cost in a distinct lower-right area.
title('Costing','B3','Unit costs and cost of goods sold','B4','Calculate unit costs, then multiply by units sold to obtain cost of goods sold.');
header('Costing','B11:H11',['Product code','Material code','Units per product','Unit price','Material per unit','Labor per unit','Unit cost']);
products.forEach((p,i)=>{
  const r=12+i,pr=9+i,m=materials.find(m=>m.code===p.mat);p.unitcost=p.use*m.unit+p.labor;p.deliveredCost=p.salesQty*p.unitcost;p.contribution=p.sales-p.deliveredCost;
  V('Costing',`B${r}`,p.sku);V('Costing',`C${r}`,p.mat);
  F('Costing',`D${r}`,`='Catalog'!L${pr}`,p.use,'Material units per product');
  F('Costing',`E${r}`,`=SUMIFS('Catalog'!$E$29:$E$31,'Catalog'!$C$29:$C$31,C${r})`,m.unit,'Unit price');
  F('Costing',`F${r}`,`=D${r}*E${r}`,p.use*m.unit,'Material per unit');
  F('Costing',`G${r}`,`='Catalog'!M${pr}`,p.labor,'Labor per unit');
  F('Costing',`H${r}`,`=SUM(F${r}:G${r})`,p.unitcost,'Unit cost');
  const r2=28+i;V('Costing',`P${r2}`,p.sku);
  F('Costing',`Q${r2}`,`='Sales'!F${pr}*H${r}`,p.deliveredCost,'Cost of units sold');
  F('Costing',`R${r2}`,`='Sales'!H${pr}-Q${r2}`,p.contribution,'Gross profit by product');
});
sheets['Costing'].getRange('D11:H15').format.columnWidth=17;
section('Costing','P25','Cost of products sold');header('Costing','P27:R27',['Product code','Month COGS','Gross profit']);
const cogs=products.reduce((a,p)=>a+p.deliveredCost,0),gross=totalSales-cogs;
V('Costing','P35','Total cost of goods sold');F('Costing','R35','=SUM(Q28:Q31)',cogs,'Total cost of goods sold');
V('Costing','P36','Total gross profit');F('Costing','R36','=SUM(R28:R31)',gross,'Total gross profit');
sheets['Costing'].getRange('P27:P36').format.columnWidth=24;sheets['Costing'].getRange('Q27:R36').format.columnWidth=19;
region('Costing','B11:H15','Product unit cost','Multiply material usage by unit price and add labor per unit.','calculation');
region('Costing','P27:R36','Cost of goods sold and gross profit','Recognize only the cost of units sold in profit.','summary');

// Production is driven by demand, opening inventory and safety inventory.
title('Inventory','J3','Finished inventory and production','J4','Calculate production needed to meet sales and the safety-stock target.');
header('Inventory','J8:P8',['Product code','Opening units','Units sold','Safety stock','Production needed','Closing units','Stock surplus']);
products.forEach((p,i)=>{
  const r=9+i,cr=12+i;p.production=Math.max(0,p.salesQty+p.safety-p.opening);p.closing=p.opening+p.production-p.salesQty;p.inventoryValue=p.closing*p.unitcost;p.productionLabor=p.production*p.labor;
  V('Inventory',`J${r}`,p.sku);
  F('Inventory',`K${r}`,`='Catalog'!O${r}`,p.opening,'Opening units');F('Inventory',`L${r}`,`='Sales'!F${r}`,p.salesQty,'Units sold');F('Inventory',`M${r}`,`='Catalog'!N${r}`,p.safety,'Target safety stock');
  F('Inventory',`N${r}`,`=MAX(0,L${r}+M${r}-K${r})`,p.production,'Production needed to fill the shortfall');
  F('Inventory',`O${r}`,`=K${r}+N${r}-L${r}`,p.closing,'Closing units');F('Inventory',`P${r}`,`=O${r}-M${r}`,p.closing-p.safety,'Difference from safety target');
  const r2=29+i;V('Inventory',`C${r2}`,p.sku);V('Inventory',`D${r2}`,p.name);
  F('Inventory',`E${r2}`,`=O${r}`,p.closing,'Closing finished units');
  F('Inventory',`F${r2}`,`=O${r}*'Costing'!H${cr}`,p.inventoryValue,'Closing finished-goods value');
  F('Inventory',`G${r2}`,`=N${r}*'Costing'!G${cr}`,p.productionLabor,'Current production labor paid');
});
sheets['Inventory'].getRange('K8:P12').format.columnWidth=17;
section('Inventory','C26','Inventory value and production labor');header('Inventory','C28:G28',['Product code','Product name','Closing units','Inventory value','Production labor']);
const inventoryValue=products.reduce((a,p)=>a+p.inventoryValue,0),labor=products.reduce((a,p)=>a+p.productionLabor,0);
V('Inventory','D36','Total');F('Inventory','F36','=SUM(F29:F32)',inventoryValue,'Total closing inventory value');F('Inventory','G36','=SUM(G29:G32)',labor,'Total production labor paid');
sheets['Inventory'].getRange('D28:D36').format.columnWidth=20;sheets['Inventory'].getRange('F28:G36').format.columnWidth=19;
region('Inventory','J8:P12','Required production calculation','Add sales and safety stock, then deduct opening finished goods to obtain required production.','calculation');
region('Inventory','C28:G36','Inventory value and labor','Closing inventory remains an asset while production labor is paid from current cash.','summary');

// Purchasing is a multi-branch join of production, BOM, inventory and unit prices.
title('Purchasing','C3','Material purchases and cash payments','C4','Deduct opening material stock from production needs and add the safety-stock target.');
section('Purchasing','O6','Materials required for production');header('Purchasing','O8:S8',['Product code','Material code','Production units','Units per product','Material units']);
products.forEach((p,i)=>{
  const r=9+i;V('Purchasing',`O${r}`,p.sku);F('Purchasing',`P${r}`,`='Catalog'!K${r}`,p.mat,'Material code for the product');F('Purchasing',`Q${r}`,`='Inventory'!N${r}`,p.production,'Production units');F('Purchasing',`R${r}`,`='Catalog'!L${r}`,p.use,'Material units per product');F('Purchasing',`S${r}`,`=Q${r}*R${r}`,p.production*p.use,'Material requirement by product');
});
sheets['Purchasing'].getRange('P8:S12').format.columnWidth=17;
section('Purchasing','C20','Purchase quantities by material');header('Purchasing','C22:K22',['Material code','Material name','Required units','Opening units','Safety units','Purchase units','Unit price','Purchase value','Cash paid']);
materials.forEach((m,i)=>{
  const r=23+i,pr=29+i;m.need=products.filter(p=>p.mat===m.code).reduce((a,p)=>a+p.production*p.use,0);m.purchase=Math.max(0,m.need+m.safety-m.opening);m.amount=m.purchase*m.unit;m.cash=m.amount*settings.pay;
  V('Purchasing',`C${r}`,m.code);V('Purchasing',`D${r}`,m.name);
  F('Purchasing',`E${r}`,`=SUMIFS($S$9:$S$12,$P$9:$P$12,C${r})`,m.need,'Required units by material');
  F('Purchasing',`F${r}`,`=SUMIFS('Assumptions'!$E$22:$E$24,'Assumptions'!$D$22:$D$24,C${r})`,m.opening,'Opening material units');
  F('Purchasing',`G${r}`,`=SUMIFS('Assumptions'!$F$22:$F$24,'Assumptions'!$D$22:$D$24,C${r})`,m.safety,'Material safety-stock target');
  F('Purchasing',`H${r}`,`=MAX(0,E${r}+G${r}-F${r})`,m.purchase,'Material purchase units');F('Purchasing',`I${r}`,`='Catalog'!E${pr}`,m.unit,'Unit price');F('Purchasing',`J${r}`,`=H${r}*I${r}`,m.amount,'Material purchase value');F('Purchasing',`K${r}`,`=J${r}*'Assumptions'!$F$10`,m.cash,'Current material cash payment');
});
const purchases=materials.reduce((a,m)=>a+m.amount,0),purchaseCash=purchases*settings.pay;
section('Purchasing','P23','Total purchases and payments');labelcol('Purchasing','P26',['Total purchase value','Cash paid','Closing payables']);
F('Purchasing','Q26','=SUM(J23:J25)',purchases,'Total purchase value');F('Purchasing','Q27','=SUM(K23:K25)',purchaseCash,'Total current material payments');F('Purchasing','Q28','=Q26-Q27',purchases-purchaseCash,'Closing trade payables');
sheets['Purchasing'].getRange('D22:D25').format.columnWidth=22;sheets['Purchasing'].getRange('E22:K25').format.columnWidth=16;sheets['Purchasing'].getRange('P26:P28').format.columnWidth=22;sheets['Purchasing'].getRange('Q26:Q28').format.columnWidth=19;
region('Purchasing','O8:S12','Material requirement by product','Multiply finished production units by material usage per product.','calculation');
region('Purchasing','C22:K25','Purchase calculation','Add safety stock to required units, deduct stock on hand, then apply the purchase price.','calculation');
region('Purchasing','P26:Q28','Total material payments','Apply the current payment rate to purchases and send the resulting cash outflow to Summary.','summary');

// Opening view: product economics upper-left, profit upper-center, cash lower-right.
title('Summary','B3','October manufacturing operations','B4','Synthetic model. Compare revenue, operating profit and closing cash. Amounts in KRW.');
section('Summary','B6','Product sales and profit');header('Summary','B7:E7',['Product code','Revenue','COGS','Gross profit']);
products.forEach((p,i)=>{
  const r=8+i,sr=9+i,cr=28+i;V('Summary',`B${r}`,p.sku);F('Summary',`C${r}`,`='Sales'!H${sr}`,p.sales,'Product revenue');F('Summary',`D${r}`,`='Costing'!Q${cr}`,p.deliveredCost,'Product cost of goods sold');F('Summary',`E${r}`,`=C${r}-D${r}`,p.contribution,'Product gross profit');
});
sheets['Summary'].getRange('C7:E11').format.columnWidth=18;
section('Summary','H6','Monthly profit');labelcol('Summary','H8',['Monthly revenue','COGS','Gross profit','Fixed operating costs','Operating profit','Operating margin']);
F('Summary','K8',"='Sales'!Q26",totalSales,'Total revenue');F('Summary','K9',"='Costing'!R35",cogs,'Cost of products sold');F('Summary','K10','=K8-K9',gross,'Gross profit');F('Summary','K11',"='Assumptions'!F11",settings.fixed,'Fixed operating costs');F('Summary','K12','=K10-K11',gross-settings.fixed,'Operating profit');F('Summary','K13','=K12/K8',(gross-settings.fixed)/totalSales,'Operating margin');
sheets['Summary'].getRange('H8:H13').format.columnWidth=21;sheets['Summary'].getRange('K8:K13').format.columnWidth=20;sheets['Summary'].getRange('K13').setNumberFormat('0.0%');
section('Summary','M20','Closing cash');labelcol('Summary','M23',['Opening cash','Current sales collected','Prior receivables collected','Materials paid','Production labor paid','Fixed costs paid','Capital expenditure paid','Closing cash','Cash above minimum']);
const endCash=settings.beginCash+cashCollect+settings.oldCollect-purchaseCash-labor-settings.fixed-settings.capex;
[
 ['P23',"='Assumptions'!F13",settings.beginCash,'Opening cash'],['P24',"='Sales'!Q27",cashCollect,'Current sales cash collections'],['P25',"='Assumptions'!F9",settings.oldCollect,'Prior receivables collected'],['P26',"='Purchasing'!Q27",purchaseCash,'Cash paid for materials'],['P27',"='Inventory'!G36",labor,'Current production labor'],['P28',"='Assumptions'!F11",settings.fixed,'Fixed operating costs paid'],['P29',"='Assumptions'!F12",settings.capex,'Capital expenditure cash payment'],['P30','=SUM(P23:P25)-SUM(P26:P29)',endCash,'Final output: closing cash'],['P31',"=P30-'Assumptions'!F14",endCash-settings.minCash,'Cash above minimum'],
].forEach(([a,f,v,l])=>F('Summary',a,f,v,l));
F('Summary','P35',`=IF(P30>='Assumptions'!F14,"Minimum cash covered","Additional cash required")`,endCash>=settings.minCash?'Minimum cash covered':'Additional cash required','Cash status');
sheets['Summary'].getRange('M23:M31').format.columnWidth=27;sheets['Summary'].getRange('P23:P35').format.columnWidth=23;
sheets['Summary'].getRange('M30:P30').format.fill='#E2F1ED';sheets['Summary'].getRange('P30').format.font={name:'Arial',size:15,bold:true,color:'#154D45'};sheets['Summary'].getRange('M30:P30').format.rowHeight=31;
section('Summary','B22','How to read the results');
['Revenue is the order value after discounts and returns.','Cost of goods sold equals units sold multiplied by unit cost.','Production covers sales and additional safety stock.','Closing cash reflects collection and payment timing and capital expenditure.'].forEach((v,i)=>V('Summary',`B${24+i}`,v));
region('Summary','B7:E11','Product sales and profit','Combine product results from Sales and Costing in one view.','output');
region('Summary','H8:K13','Monthly profit','Deduct cost of goods sold and fixed operating costs from revenue.','output');
region('Summary','M23:P35','Closing cash calculation','Combine inflows and outflows to calculate closing cash in P30.','output');

// Restrained totals and formats, with no full-sheet grid or decorative drawings.
for(const [s,r] of [['Sales','P25:Q29'],['Costing','P35:R36'],['Inventory','D36:G36'],['Purchasing','P26:Q28'],['Summary','H12:K13']]){
  sheets[s].getRange(r).format.font.bold=true;
  sheets[s].getRange(r).format.borders={top:{style:'thin',color:'#A8BBC5'}};
}
wb.recalculate();
const spec={
  version:1,title:'October operations at a fictional manufacturer',synthetic:true,
  description:'An eight-sheet synthetic model linking sales, production, purchasing, costs and cash for four products, three materials and sixteen orders.',
  workbookFile:'OperationsModel.xlsx',sheetOrder:sheetNames,regions,cells,outputs:expected,
  headline:{sheet:'Summary',address:'P30',title:'Closing cash',expected:endCash},
  metrics:{formulas:expected.length,orders:orders.length,products:products.length,materials:materials.length,totalSales,cogs,gross,profit:gross-settings.fixed,cashCollect,purchases,purchaseCash,labor,inventoryValue,endCash},
  demo:{title:'Where does closing cash come from?',description:'Select the result to reveal sales collections, production and purchasing outflows at their actual worksheet positions.',startSheet:'Summary',startAddress:'P30',stops:[
    {sheet:'Summary',address:'P30',title:'Start here',body:'Closing cash is opening cash plus collections, less materials, labor, fixed costs and capital expenditure.'},
    {sheet:'Sales',address:'Q27',title:'Cash coming in',body:'Apply the 80% collection rate in Assumptions to each product’s monthly sales.'},
    {sheet:'Orders',address:'K8',title:'Where revenue starts',body:'Apply Catalog prices and each order’s discount to units sold after returns.'},
    {sheet:'Inventory',address:'N9',title:'Orders become production',body:'Add units sold and safety stock, then deduct opening finished goods to calculate production.'},
    {sheet:'Purchasing',address:'S9',title:'Production becomes materials',body:'Multiply production by each product’s material usage. Purchasing combines products that use the same material.'},
    {sheet:'Purchasing',address:'H23',title:'Purchases after inventory',body:'Add the material safety target and deduct opening stock to calculate new purchases.'},
    {sheet:'Purchasing',address:'Q27',title:'Purchases become cash outflows',body:'Only 65% of purchase units times unit price leaves cash this month.'},
    {sheet:'Inventory',address:'G36',title:'Another cash outflow',body:'Sum production units times labor per unit. Cost of goods sold and production labor payments measure different things.'},
    {sheet:'Summary',address:'P30',title:'The complete calculation',body:'Sales collections, material purchases and production labor meet in one result. The flow explains why profit and cash differ.'},
  ]},
  perturbation:{sheet:'Assumptions',address:'F8',base:settings.collect,test:.9,outputSheet:'Summary',outputAddress:'P30',expected:endCash+totalSales*.1},
};
await fs.writeFile(path.join(dir,'model-spec.json'),JSON.stringify(spec,null,2)+'\n');
const formulaChecks = expected.map(item => {
  const actual = sheets[item.sheet].getRange(item.address).values[0][0];
  const pass = typeof item.expected === 'number'
    ? typeof actual === 'number' && Math.abs(actual - item.expected) <= Math.max(0.000001, Math.abs(item.expected) * 1e-10)
    : actual === item.expected;
  return { sheet: item.sheet, address: item.address, expected: item.expected, actual, pass };
});
const baseProfit = sheets.Summary.getRange('K12').values[0][0];
sheets.Assumptions.getRange('F8').values = [[0.9]];
wb.recalculate();
const changedCash = sheets.Summary.getRange('P30').values[0][0];
const changedCollections = sheets.Sales.getRange('Q27').values[0][0];
const changedProfit = sheets.Summary.getRange('K12').values[0][0];
const perturbation = {
  input: 'Assumptions!F8', base: 0.8, test: 0.9,
  output: 'Summary!P30', actual: changedCash, expected: endCash + totalSales * 0.1,
  collections: changedCollections, expectedCollections: totalSales * 0.9,
  profitUnchanged: changedProfit === baseProfit,
  pass: Math.abs(changedCash - (endCash + totalSales * 0.1)) < 0.000001
    && Math.abs(changedCollections - totalSales * 0.9) < 0.000001 && changedProfit === baseProfit
};
sheets.Assumptions.getRange('F8').values = [[settings.collect]];
wb.recalculate();
const restored = Math.abs(sheets.Summary.getRange('P30').values[0][0] - endCash) < 0.000001
  && sheets.Assumptions.getRange('F8').values[0][0] === settings.collect;
const modelChecks = {
  engine: '@oai/artifact-tool', nativeExcelRecalculationVerified: false,
  formulaChecks, perturbation, restored,
  passed: formulaChecks.filter(item => item.pass).length,
  failed: formulaChecks.filter(item => !item.pass).length
};
await fs.writeFile(path.join(qaDir, 'model_checks.json'), JSON.stringify(modelChecks, null, 2));
if (modelChecks.failed || !perturbation.pass || !restored) throw new Error('Model calculation checks failed. See model_checks.json.');
console.log(`Calculated and checked ${modelChecks.passed} formulas, collection-rate sensitivity and restored base inputs.`);
const xlsx=await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(path.join(dir, 'OperationsModel.xlsx'));
try { await fs.rename(path.join(dir, 'OperationsModel.xlsx.inspect.ndjson'), path.join(qaDir, 'export_inspect.ndjson')); }
catch (error) { if (error.code !== 'ENOENT') throw error; }
await fs.copyFile(path.join(dir, 'OperationsModel.xlsx'), path.join(repoDir, 'docs', 'downloads', 'OperationsModel.xlsx'));
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:20},maxChars:4000});
await fs.writeFile(path.join(qaDir,'example_formula_scan.ndjson'),errors.ndjson);
for(const s of sheetNames){
  console.log(`Rendering ${s}.`);
  const preview=await wb.render({sheetName:s,range:s==='Orders'?'A1:L25':'A1:U38',scale:1,format:'png'});
  await fs.writeFile(path.join(qaDir,`example_preview_${s}.png`),new Uint8Array(await preview.arrayBuffer()));
}
const summary=await wb.inspect({kind:'table',range:'Summary!M23:P31',include:'values,formulas',tableMaxRows:9,tableMaxCols:4,maxChars:5000});
await fs.writeFile(path.join(qaDir,'example_inspect.ndjson'),summary.ndjson);
console.log(JSON.stringify({file:'examples/OperationsModel.xlsx',formulas:expected.length,metrics:spec.metrics,formulaChecks:modelChecks.passed,perturbationPassed:perturbation.pass,restored},null,2));
