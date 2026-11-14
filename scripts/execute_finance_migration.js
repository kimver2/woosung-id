const fs = require('fs');
const path = require('path');

const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzd6pJp0oDfhN55dv6xM1HsAoQikA2Y_ApMz-H-GpaYqf3EKh3ffh6ULz-bU9n24dfTbg/exec';

async function main() {
  const payloadPath = path.join(__dirname, '..', 'data', 'finance_migration_payload.json');
  console.log('Loading migration payload from:', payloadPath);
  const rawData = fs.readFileSync(payloadPath, 'utf8');
  const payload = JSON.parse(rawData);

  console.log('Payload summary:');
  console.log('  rawCollection:', payload.rawCollection.length, 'rows');
  console.log('  rawExpense:', payload.rawExpense.length, 'rows');
  console.log('  collectionRows:', payload.collectionRows.length, 'rows');
  console.log('  expenseRows:', payload.expenseRows.length, 'rows');
  console.log('  fixedMasterRows:', payload.fixedMasterRows.length, 'rows');

  console.log('\nSending setupFinanceDbAndMigrate POST request to GAS Web App...');
  const res = await fetch(WEB_APP_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      action: 'setupFinanceDbAndMigrate',
      data: payload
    }),
    redirect: 'follow'
  });

  const resText = await res.text();
  console.log('Response status:', res.status);
  try {
    const json = JSON.parse(resText);
    console.log('Response JSON:', JSON.stringify(json, null, 2));
    if (json.success) {
      console.log('\n🎉 재무 DB 5대 시트 초기 구축 및 데이터 마이그레이션 성공!');
    } else {
      console.error('\n❌ 마이그레이션 실패:', json.error);
      process.exit(1);
    }
  } catch (err) {
    console.log('Response text (not JSON):', resText);
  }
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
