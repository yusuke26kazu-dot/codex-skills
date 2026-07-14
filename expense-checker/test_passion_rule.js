// Test script to verify Passion Leaders logic
function validateRow(row) {
    const issues = [];
    function addIssue(title, desc) {
        issues.push({ title, desc });
    }

    // 1.1 Project ID check
    if (!row.projId) {
        addIssue('プロジェクトID未入力', 'プロジェクトIDが入力されていません。');
    } else {
        const hasValidProjId = row.projId.includes('電子雑誌') || 
                              row.projId.includes('飲食') || 
                              row.projId.includes('パッションリーダーズ');
        if (!hasValidProjId) {
            addIssue('プロジェクトID不正', 'プロジェクトIDが正しくありません。');
        } else {
            const hasPassionId = row.projId.includes('パッションリーダーズ');
            const hasPassionPurpose = row.remarks.includes('利用目的【パッションリーダーズ手伝いのため】');
            if (hasPassionId && !hasPassionPurpose) {
                addIssue('プロジェクトID不整合', 'プロジェクトIDが「パッションリーダーズ」ですが、備考欄に「利用目的【パッションリーダーズ手伝いのため】」が記載されていません。');
            } else if (!hasPassionId && hasPassionPurpose) {
                addIssue('プロジェクトID不整合', '備考欄に「利用目的【パッションリーダーズ手伝いのため】」と記載されていますが、プロジェクトIDが「パッションリーダーズ」になっていません。');
            }
        }
    }

    // 2.1 Normal transport purpose check
    const isNormalTransport = row.category.startsWith('交通費') && 
                              !row.category.includes('駐車場') && 
                              !row.category.includes('ガソリン');
    if (isNormalTransport) {
        const isPassionLeaders = row.projId && row.projId.includes('パッションリーダーズ') && row.remarks.includes('利用目的【パッションリーダーズ手伝いのため】');
        const hasPassionPurpose = row.remarks.includes('利用目的【パッションリーダーズ手伝いのため】');
        if (!isPassionLeaders && !hasPassionPurpose) {
            if (!row.remarks.includes('【旅色営業のため】')) {
                addIssue('利用目的不備', '交通費の備考欄に利用目的「【旅色営業のため】」が含まれていません。');
            }
        }
    }

    return issues;
}

const testCases = [
    {
        name: 'Normal Tabiiro Sales OK',
        row: { category: '交通費(電車)', projId: '15(飲食)', remarks: '利用目的【旅色営業のため】' },
        expected: []
    },
    {
        name: 'Passion Leaders OK',
        row: { category: '交通費(電車)', projId: '17(パッションリーダーズ)', remarks: '利用目的【パッションリーダーズ手伝いのため】' },
        expected: []
    },
    {
        name: 'Tabiiro Sales Missing Purpose',
        row: { category: '交通費(電車)', projId: '15(飲食)', remarks: '別の目的' },
        expected: [{ title: '利用目的不備' }]
    },
    {
        name: 'Passion Leaders Missing Purpose',
        row: { category: '交通費(電車)', projId: '17(パッションリーダーズ)', remarks: '利用目的【旅色営業のため】' },
        expected: [{ title: 'プロジェクトID不整合' }]
    },
    {
        name: 'Passion Leaders Purpose on Tabiiro ID',
        row: { category: '交通費(電車)', projId: '15(飲食)', remarks: '利用目的【パッションリーダーズ手伝いのため】' },
        expected: [{ title: 'プロジェクトID不整合' }]
    },
    {
        name: 'Invalid Project ID',
        row: { category: '交通費(電車)', projId: '18(その他)', remarks: '利用目的【旅色営業のため】' },
        expected: [{ title: 'プロジェクトID不正' }]
    }
];

testCases.forEach(tc => {
    const res = validateRow(tc.row);
    const pass = JSON.stringify(res.map(i => i.title)) === JSON.stringify(tc.expected.map(i => i.title));
    console.log(`Test [${tc.name}]: ${pass ? 'PASS' : 'FAIL'}`);
    if (!pass) {
        console.log(`  Result:   `, res);
        console.log(`  Expected: `, tc.expected);
    }
});
