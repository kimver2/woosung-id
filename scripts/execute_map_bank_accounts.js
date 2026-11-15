const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzd6pJp0oDfhN55dv6xM1HsAoQikA2Y_ApMz-H-GpaYqf3EKh3ffh6ULz-bU9n24dfTbg/exec';

async function main() {
  console.log('Sending mapPartnerBankAccounts request to GAS Web App...');
  const res = await fetch(WEB_APP_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      action: 'mapPartnerBankAccounts'
    }),
    redirect: 'follow'
  });

  const resText = await res.text();
  console.log('Response status:', res.status);
  try {
    const json = JSON.parse(resText);
    console.log('Response JSON:', JSON.stringify(json, null, 2));
    if (json.success) {
      console.log('\n🎉 거래처 계좌번호 일괄 매핑 성공!');
      console.log('  - 매칭 건수:', json.matchedCount);
      console.log('  - 업데이트 건수:', json.updatedCount);
      console.log('  - 총 거래처 수:', json.totalPartners);
      console.log('  - 매핑 샘플:', json.samples);
      if (json.unmatchedPartners && json.unmatchedPartners.length > 0) {
        console.log('  - 미매칭 업체 (총 ' + json.unmatchedPartners.length + '개 중 상위 10개):', json.unmatchedPartners.slice(0, 10));
      }
    } else {
      console.error('\n❌ 매핑 실패:', json.error);
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
