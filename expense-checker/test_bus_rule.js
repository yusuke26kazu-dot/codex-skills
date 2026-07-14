// Test script to verify the bus-only detection regex and logic
const busCompanyPattern = /(東急|ＪＲ|JR|小田急|京王|西武|東武|京急|京成|名鉄|近鉄|南海|阪急|都営|都|市営|市|関東|国際興業|川崎鶴見臨港|立川|相鉄|相模鉄道|コミュニティ|シャトル)バス/g;

const trainKeywords = [
    'ＪＲ', 'JR', 'メトロ', '地下鉄', '小田急', '京王', '東急', '西武', '東武', '京急', '京成', '相鉄', 
    'つくばエクスプレス', '新幹線', 'モノレール', '線', '電鉄', '鉄道', '急行', '快速', '特急'
];

function isBusOnly(payee, remarks = '') {
    const hasBusText = payee.includes('バス') || remarks.includes('バス');
    if (!hasBusText) return false;

    let isBusOnlyTrip = false;
    const hasArrows = payee.includes('→') || payee.includes('⇔');
    if (hasArrows) {
        const parts = payee.split(/→|⇔/);
        if (parts.length >= 2) {
            const startHasBus = parts[0].includes('バス');
            const endHasBus = parts[parts.length - 1].includes('バス');
            if (startHasBus && endHasBus) {
                isBusOnlyTrip = true;
            }
        }
    } else {
        isBusOnlyTrip = true;
    }

    if (!isBusOnlyTrip) return false;

    const normalizedPayee = payee.replace(busCompanyPattern, 'バス');
    const normalizedRemarks = remarks.replace(busCompanyPattern, 'バス');

    const hasTrainKeyword = trainKeywords.some(kw => normalizedPayee.includes(kw) || normalizedRemarks.includes(kw));
    return !hasTrainKeyword;
}

const testCases = [
    { payee: '渋谷駅(バス)→目黒郵便局(バス)', remarks: '', expected: true },
    { payee: '渋谷駅→目黒郵便局(バス)', remarks: '', expected: false }, // only end has bus, so "train -> bus" -> false
    { payee: '渋谷駅(東急バス)→目黒郵便局(東急バス)', remarks: '', expected: true },
    { payee: '渋谷駅→目黒郵便局(都営バス)', remarks: '', expected: false }, // only end has bus -> false
    { payee: '品川駅(JR)→学芸大学(東急)→野沢三丁目(バス)', remarks: '', expected: false },
    { payee: '本町→堺筋本町(地下鉄)', remarks: 'バスで乗り継ぎ', expected: false },
    { payee: 'JRバス関東 乗り場', remarks: '', expected: true }, // no arrows -> true
    { payee: '東急バス乗車', remarks: '', expected: true }, // no arrows -> true
    { payee: '本町(バス)⇔渋谷(バス)', remarks: '東名ハイウェイバス', expected: true },
    { payee: '渋谷駅→上原二丁目(バス)', remarks: '【旅色営業のため】', expected: false } // only end has bus -> false
];

testCases.forEach((tc, i) => {
    const result = isBusOnly(tc.payee, tc.remarks);
    console.log(`Test ${i + 1}: Payee="${tc.payee}", Remarks="${tc.remarks}"`);
    console.log(`  Result: ${result} (Expected: ${tc.expected}) - ${result === tc.expected ? 'PASS' : 'FAIL'}`);
});
