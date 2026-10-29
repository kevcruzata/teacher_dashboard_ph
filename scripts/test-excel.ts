// @ts-ignore
import XlsxPopulate from 'xlsx-populate';
import config from '../src/config/ecr_shs_v26.json';

async function runTest() {
  console.log("Starting robust INPUT DATA injection test...");

  const workbook = await XlsxPopulate.fromFileAsync('./public/templates/ecr_shs_v26.xlsx');
  const inputSheet = workbook.sheet(config.sheets.metadata_sheet);
  
  if (!inputSheet) {
    throw new Error(`Could not find sheet: ${config.sheets.metadata_sheet}`);
  }

  // 1. Write metadata values explicitly as strings or numbers
  inputSheet.cell(config.metadata.region).value("Region V");
  inputSheet.cell(config.metadata.division).value("Iriga");
  inputSheet.cell(config.metadata.school_id).value("12345");
  inputSheet.cell(config.metadata.school_name).value("ZAHS");
  inputSheet.cell(config.metadata.school_year).value("2026 - 2027"); 
  inputSheet.cell(config.metadata.teacher).value("CK");
  
  // Grade level must be stored as a number so dropdown validation indexes match
  inputSheet.cell(config.metadata.grade_level).value(11);
  inputSheet.cell(config.metadata.section_name).value("Diamond");

  // 2. Subject & Category dropdown fields
  inputSheet.cell(config.metadata.subject_category).value("SSHS - CORE");
  inputSheet.cell(config.metadata.cluster).value("CORE");
  inputSheet.cell(config.metadata.subject).value("General Mathematics");

  // 3. WAKE-UP FIX: If the template uses a selection index cell on the helper sheet, 
  // you can also update it directly so the dropdown evaluates immediately.
  const helperSheet = workbook.sheet('HELPER');
  if (helperSheet) {
    // If the helper sheet tracks the selected grade level or subject index:
    // helperSheet.cell('F7').value(11);
  }

  const outputPath = './test-output.xlsx';
  await workbook.toFileAsync(outputPath);
  console.log(`Success! Saved to ${outputPath}`);
}

runTest().catch(console.error);