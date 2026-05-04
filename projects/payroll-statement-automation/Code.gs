/**
 * 급여명세서 자동 생성 스크립트 (Google Apps Script)
 *
 * 사용 방법
 *  1) 구글 시트에서 [확장 프로그램] > [Apps Script] 열기
 *  2) 이 파일 전체 코드를 붙여넣고 저장
 *  3) 시트를 새로고침하면 상단 메뉴에 [💰 급여명세서] 가 나타남
 *  4) [🛠 양식 시트 만들기/초기화] 1회 실행 → "급여명세서_양식" 탭이 자동 생성됨
 *  5) [▶ PDF 일괄 생성] 클릭 → 귀속 연월 입력 → 직원별 PDF가 드라이브 폴더에 저장됨
 */

// ===================== 설정값 =====================
const CONFIG = {
  PAYROLL_SHEET_NAME: '이번달 급여정보',                       // 급여대장 시트 탭 이름
  TEMPLATE_SHEET_NAME: '급여명세서_양식',                      // 양식 시트 탭 이름 (자동 생성)
  DRIVE_FOLDER_ID: '1ahK5JPeczmFp8NlYkyXq7_KNBHNzl7_B',       // PDF 저장 드라이브 폴더 ID
  COMPANY_NAME: '주식회사 파크홈',                              // 회사명 (수정해서 사용)
  TIMEZONE: 'Asia/Seoul',
};

// 급여대장 시트의 컬럼 위치 (1-base, A=1)
const COL = {
  순번: 1, 이름: 2, 직책: 3, 입사일: 4, 부양가족: 5,
  기본급: 6, 상여: 7, 식대: 8, 자가운전보조금: 9, 육아수당: 10, 연구보조금: 11,
  과세액계: 12, 지급액계: 13,
  국민연금: 14, 건강보험: 15, 고용보험: 16, 장기요양보험료: 17,
  소득세: 18, 지방소득세: 19,
  기타차감1: 20, 기타차감2: 21, 기타차감3: 22, 기타차감4: 23,
  공제액계: 24, 차인지급액: 25,
};

// ===================== 메뉴 등록 =====================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 급여명세서')
    .addItem('▶ PDF 일괄 생성', 'generatePayslipsPDF')
    .addSeparator()
    .addItem('🛠 양식 시트 만들기/초기화', 'createTemplateSheet')
    .addToUi();
}

// ===================== 메인: PDF 일괄 생성 =====================
function generatePayslipsPDF() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const defaultYM = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM');
  const response = ui.prompt(
    '귀속 연월 입력',
    `급여명세서 귀속 연월을 입력하세요. (예: ${defaultYM})`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const yearMonth = (response.getResponseText() || '').trim() || defaultYM;
  if (!/^\d{4}-\d{2}$/.test(yearMonth)) {
    ui.alert('형식 오류', 'YYYY-MM 형식으로 입력해주세요. (예: 2026-05)', ui.ButtonSet.OK);
    return;
  }

  const payrollSheet = ss.getSheetByName(CONFIG.PAYROLL_SHEET_NAME);
  if (!payrollSheet) {
    ui.alert('오류', `'${CONFIG.PAYROLL_SHEET_NAME}' 시트를 찾을 수 없습니다.`, ui.ButtonSet.OK);
    return;
  }

  let templateSheet = ss.getSheetByName(CONFIG.TEMPLATE_SHEET_NAME);
  if (!templateSheet) templateSheet = createTemplateSheetInternal(ss);

  const lastRow = payrollSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('오류', '급여 데이터가 없습니다.', ui.ButtonSet.OK);
    return;
  }
  const data = payrollSheet.getRange(2, 1, lastRow - 1, 25).getValues();

  let folder;
  try {
    folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  } catch (e) {
    ui.alert('오류', '드라이브 폴더에 접근할 수 없습니다. 폴더 ID와 권한을 확인하세요.', ui.ButtonSet.OK);
    return;
  }

  const monthFolder = getOrCreateSubFolder(folder, yearMonth);

  let count = 0;
  const failed = [];
  for (const row of data) {
    const name = row[COL.이름 - 1];
    if (!name || String(name).trim() === '') continue;

    try {
      fillTemplate(templateSheet, row, yearMonth);
      SpreadsheetApp.flush();
      const pdfBlob = exportSheetToPDF(ss, templateSheet);
      pdfBlob.setName(`${yearMonth}_${name}_급여명세서.pdf`);
      monthFolder.createFile(pdfBlob);
      count++;
    } catch (err) {
      failed.push(`${name}: ${err.message}`);
    }
  }

  let msg = `${count}명의 급여명세서를 생성했습니다.\n저장 위치: ${folder.getName()} / ${yearMonth}`;
  if (failed.length) msg += `\n\n실패 ${failed.length}건:\n` + failed.join('\n');
  ui.alert('완료', msg, ui.ButtonSet.OK);
}

// ===================== 양식에 데이터 채우기 =====================
function fillTemplate(sheet, row, yearMonth) {
  const v = (i) => row[i - 1];

  sheet.getRange('B2').setValue(`${yearMonth} 급여명세서`);

  // 인적사항
  sheet.getRange('C4').setValue(v(COL.순번));
  sheet.getRange('F4').setValue(CONFIG.COMPANY_NAME);
  sheet.getRange('C5').setValue(v(COL.이름));
  sheet.getRange('C6').setValue(v(COL.직책) || '');
  sheet.getRange('C7').setValue(formatDate(v(COL.입사일)));

  // 지급
  sheet.getRange('C10').setValue(v(COL.기본급) || 0);
  sheet.getRange('C11').setValue(v(COL.상여) || 0);
  sheet.getRange('C12').setValue(v(COL.식대) || 0);
  sheet.getRange('C13').setValue(v(COL.자가운전보조금) || 0);
  sheet.getRange('C14').setValue(v(COL.육아수당) || 0);
  sheet.getRange('C15').setValue(v(COL.연구보조금) || 0);
  sheet.getRange('C16').setValue(v(COL.지급액계) || 0);

  // 공제
  sheet.getRange('F10').setValue(v(COL.국민연금) || 0);
  sheet.getRange('F11').setValue(v(COL.건강보험) || 0);
  sheet.getRange('F12').setValue(v(COL.장기요양보험료) || 0);
  sheet.getRange('F13').setValue(v(COL.고용보험) || 0);
  sheet.getRange('F14').setValue(v(COL.소득세) || 0);
  sheet.getRange('F15').setValue(v(COL.지방소득세) || 0);

  const etc = num(v(COL.기타차감1)) + num(v(COL.기타차감2)) + num(v(COL.기타차감3)) + num(v(COL.기타차감4));
  sheet.getRange('F16').setValue(etc);
  sheet.getRange('F17').setValue(v(COL.공제액계) || 0);

  // 차인지급액
  sheet.getRange('C19').setValue(v(COL.차인지급액) || 0);

  // 발행일
  sheet.getRange('B20:F20').merge()
    .setValue(`발행일: ${Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd')}     |     ${CONFIG.COMPANY_NAME} (인)`)
    .setHorizontalAlignment('right');
}

// ===================== 양식 시트 생성 =====================
function createTemplateSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName(CONFIG.TEMPLATE_SHEET_NAME);
  if (existing) {
    const res = ui.alert(
      '양식 시트 초기화',
      '기존 양식 시트가 있습니다. 새로 만들면 기존 양식이 삭제됩니다. 계속하시겠습니까?',
      ui.ButtonSet.YES_NO
    );
    if (res !== ui.Button.YES) return;
    ss.deleteSheet(existing);
  }
  createTemplateSheetInternal(ss);
  ui.alert('완료', `'${CONFIG.TEMPLATE_SHEET_NAME}' 시트가 생성되었습니다.`, ui.ButtonSet.OK);
}

function createTemplateSheetInternal(ss) {
  const sheet = ss.insertSheet(CONFIG.TEMPLATE_SHEET_NAME);

  // 컬럼 너비
  [30, 110, 160, 30, 110, 160, 30].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // 제목
  sheet.getRange('B2:F2').merge()
    .setValue('YYYY-MM 급여명세서')
    .setFontSize(20).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#f5f5f5');
  sheet.setRowHeight(2, 44);

  // 인적사항
  const labelStyle = (range) => range.setFontWeight('bold').setBackground('#e8f0fe').setHorizontalAlignment('center').setVerticalAlignment('middle');
  labelStyle(sheet.getRange('B4')).setValue('사원번호');
  labelStyle(sheet.getRange('E4')).setValue('회 사 명');
  labelStyle(sheet.getRange('B5')).setValue('성    명');
  labelStyle(sheet.getRange('B6')).setValue('직    책');
  labelStyle(sheet.getRange('B7')).setValue('입 사 일');
  sheet.getRange('E5:F7').merge();
  for (let r = 4; r <= 7; r++) sheet.setRowHeight(r, 26);

  // 지급/공제 섹션 헤더
  sheet.getRange('B9:C9').merge().setValue('지 급 내 역')
    .setFontWeight('bold').setBackground('#cfe3cf').setHorizontalAlignment('center');
  sheet.getRange('E9:F9').merge().setValue('공 제 내 역')
    .setFontWeight('bold').setBackground('#f0cfcf').setHorizontalAlignment('center');
  sheet.setRowHeight(9, 28);

  // 지급 항목
  const incomeLabels = ['기본급', '상여', '식대', '자가운전보조금', '육아수당', '연구보조금', '지급합계'];
  for (let i = 0; i < incomeLabels.length; i++) {
    const r = 10 + i;
    sheet.getRange(r, 2).setValue(incomeLabels[i]).setBackground('#f3faf3').setHorizontalAlignment('left').setFontWeight(i === 6 ? 'bold' : 'normal');
    sheet.getRange(r, 3).setNumberFormat('#,##0').setHorizontalAlignment('right').setFontWeight(i === 6 ? 'bold' : 'normal');
    if (i === 6) {
      sheet.getRange(r, 2, 1, 2).setBackground('#cfe3cf');
    }
    sheet.setRowHeight(r, 24);
  }

  // 공제 항목
  const deductLabels = ['국민연금', '건강보험', '장기요양보험료', '고용보험', '소득세', '지방소득세', '기타차감', '공제합계'];
  for (let i = 0; i < deductLabels.length; i++) {
    const r = 10 + i;
    sheet.getRange(r, 5).setValue(deductLabels[i]).setBackground('#fdf3f3').setHorizontalAlignment('left').setFontWeight(i === 7 ? 'bold' : 'normal');
    sheet.getRange(r, 6).setNumberFormat('#,##0').setHorizontalAlignment('right').setFontWeight(i === 7 ? 'bold' : 'normal');
    if (i === 7) {
      sheet.getRange(r, 5, 1, 2).setBackground('#f0cfcf');
    }
    sheet.setRowHeight(r, 24);
  }

  // 실수령액
  sheet.getRange('B19').setValue('실 수 령 액').setFontWeight('bold').setBackground('#fff2cc')
    .setFontSize(13).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('C19').setNumberFormat('#,##0').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('right').setVerticalAlignment('middle').setBackground('#fff8e0');
  sheet.getRange('E19:F19').merge().setValue('단위: 원')
    .setHorizontalAlignment('right').setBackground('#fff2cc').setVerticalAlignment('middle');
  sheet.setRowHeight(19, 36);

  // 발행일 줄
  sheet.getRange('B20:F20').merge();
  sheet.setRowHeight(20, 30);

  // 테두리
  sheet.getRange('B2:F19').setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('B2:F2').setBorder(true, true, true, true, false, false, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 격자/탭 정리
  sheet.setHiddenGridlines(true);
  sheet.setTabColor('#fbbc04');

  return sheet;
}

// ===================== PDF 익스포트 =====================
function exportSheetToPDF(ss, sheet) {
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?` + [
    'exportFormat=pdf',
    'format=pdf',
    'size=A4',
    'portrait=true',
    'fitw=true',
    'top_margin=0.5',
    'bottom_margin=0.5',
    'left_margin=0.5',
    'right_margin=0.5',
    'sheetnames=false',
    'printtitle=false',
    'pagenumbers=false',
    'gridlines=false',
    'fzr=false',
    `gid=${sheet.getSheetId()}`,
  ].join('&');

  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
  });
  return res.getBlob();
}

// ===================== 유틸 =====================
function getOrCreateSubFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function formatDate(d) {
  if (!d) return '';
  if (d instanceof Date) return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  return String(d);
}

function num(x) {
  const n = Number(x);
  return isNaN(n) ? 0 : n;
}
