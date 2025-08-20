// 검증자 선정 함수
function selectVerifiers() {
  if (!hasUsers()) {
    console.log('⚠️ clickDB.xlsx에 사용자가 없습니다. 검증자를 선정할 수 없습니다.');
    return [];
  }

  try {
    const wb = XLSX.readFile(CONFIRM_SCORE_PATH);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    // 첫 행은 헤더이므로 제외
    const rows = data.slice(1);

    // 멤버 객체 생성
    const members = rows.map(row => ({
      id: row ? row.toString().trim() : '',
      score: row ? parseFloat(row) : 0
    })).filter(m => m.id); // id가 없는 경우 제외

    // 멤버 수
    const n = members.length;

    // 점수 내림차순 > id 오름차순 정렬
    members.sort((a, b) => {
      if (b.score !== a.score) return b.score - a.score;
      return a.id.localeCompare(b.id);
    });

    // 검증자 수 결정
    let verifierCount = 0;
    if (n < 4) verifierCount = n;
    else if (n <= 10) verifierCount = 3;
    else if (n <= 99) verifierCount = 5;
    else verifierCount = 10;

    // 조건(0.5 이상)에 맞는 후보만 선정
    const candidates = members.filter(m => m.score >= 0.5);
    const verifiers = candidates.slice(0, verifierCount);

    console.log('=== 검증자 선정 결과 ===');
    if (verifiers.length === 0) {
      console.log('⚠️ 조건(0.5 이상)에 맞는 검증자가 없습니다.');
    } else {
      verifiers.forEach((v, idx) => {
        console.log(`${idx + 1}. ${v.id} (점수: ${v.score})`);
      });
    }

    return verifiers;
  } catch (err) {
    console.error('Error reading ConfirmScoreDB.xlsx:', err);
    return [];
  }
}
