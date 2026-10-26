/** PTTB Demo — only the synthetic exam workbook is permitted. */
const DEMO_SHEET_ID = '1y5jLreJ3prUT1GGbECVhXkxpG2CUn67NeWxZXiRT7nE';
function onOpen() {
  SpreadsheetApp.getUi().createMenu('PTTB Demo')
    .addItem('Mở Dashboard', 'showPTTBDashboard')
    .addItem('Kiểm tra dữ liệu', 'testPTTBDashboard').addToUi();
}
function showPTTBDashboard() {
  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutputFromFile('DashboardPTTB').setWidth(1180).setHeight(760),
    'PTTB • Dashboard minh họa');
}
function doGet() {
  return HtmlService.createHtmlOutputFromFile('DashboardPTTB').setTitle('PTTB Demo');
}
function getPTTBDashboardData() {
  const ss = SpreadsheetApp.openById(DEMO_SHEET_ID);
  SpreadsheetApp.flush();
  const source = ss.getSheetByName('DATA_MASTER');
  const raw = source.getRange(2, 1, Math.max(1, source.getLastRow()-1), 8).getValues();
  const seen = new Set();
  raw.forEach(function(r) {
    if (!r[0]) return;
    if (!/^DEMO-HD-\d{3}$/.test(String(r[0])) || !/^Khách hàng mẫu \d+$/.test(String(r[1])) ||
        !/^DEMO-NV-\d{2}$/.test(String(r[4])) || !/^Đội mẫu [ABC]$/.test(String(r[5])))
      throw new Error('Chỉ cho phép dữ liệu giả lập trong bản bài thi.');
    if (seen.has(r[0])) throw new Error('Trùng hợp đồng mẫu: '+r[0]);
    seen.add(r[0]);
  });
  const report = ss.getSheetByName('BAO_CAO_THANG');
  const period = report.getRange('A3:E3').getValues()[0];
  const month = Number(period[1]), year = Number(period[4]);
  if (!Number.isInteger(month) || month<1 || month>12 || !Number.isInteger(year))
    throw new Error('Kiểm tra tháng tại B3 và năm tại E3.');
  const staff = report.getRange('A6:H17').getValues().filter(r=>r[0]).map(r=>({
    id:String(r[0]),name:String(r[1]),team:String(r[2]),plan:Number(r[3])||0,
    online:Number(r[4])||0,pending:Number(r[5])||0,rate:Number(r[6])||0,cumulative:Number(r[7])||0
  }));
  if (!staff.every(r=>/^DEMO-NV-\d{2}$/.test(r.id) && /^Nhân sự mẫu \d{2}$/.test(r.name)))
    throw new Error('Nhân sự phải là dữ liệu mẫu.');
  const totals = report.getRange('D19:H19').getValues()[0];
  const summary={plan:Number(totals[0])||0,online:Number(totals[1])||0,pending:Number(totals[2])||0,
    rate:Number(totals[3])||0,cumulative:Number(totals[4])||0,active:staff.filter(r=>r.online>0).length};
  ['plan','online','pending','cumulative'].forEach(k=>{
    if(staff.reduce((a,r)=>a+r[k],0)!==summary[k]) throw new Error('Đối soát không khớp: '+k);
  });
  const teams = [...new Set(staff.map(r=>r.team))].sort().map(team=>{
    const list=staff.filter(r=>r.team===team);
    const plan=list.reduce((a,r)=>a+r.plan,0), online=list.reduce((a,r)=>a+r.online,0);
    return {team,plan,online,rate:plan?online/plan:0,pending:list.reduce((a,r)=>a+r.pending,0)};
  });
  return {month,year,summary,staff,teams,synthetic:true,contractCount:seen.size};
}
function testPTTBDashboard() {
  const data=getPTTBDashboardData();
  console.log(JSON.stringify({ok:true,month:data.month,year:data.year,contracts:data.contractCount,
    staff:data.staff.length,summary:data.summary}));
  return data.summary;
}
