// =============================================================================
// sog-to-js.js  —  .sog 파일을 base64로 인코딩해 HTML에 내장 가능한 .js로 변환
//
// 용도: 웹서버 없이 HTML을 더블클릭(file://)만 해도 3DGS 모델이 보이도록
//       .sog 데이터를 window.__SOG_B64__ 전역 변수로 박아 넣은 .js를 생성합니다.
//
// 사용법 (Node.js 설치 필요):
//   node sog-to-js.js <입력.sog> [출력.js]
//
// 예시:
//   node sog-to-js.js assets/물문화관_3DGS_SPLAT.sog
//     → assets/물문화관_3DGS_SPLAT_sog.js 자동 생성 (출력 경로 생략 시)
//
//   node sog-to-js.js assets/장미공원.sog assets/장미공원sog.js
//     → 출력 파일명을 직접 지정
// =============================================================================

const fs = require('fs');
const path = require('path');

// ---- 1) 명령행 인수 파싱 -----------------------------------------------------
const args = process.argv.slice(2);

if (args.length < 1 || args[0] === '-h' || args[0] === '--help') {
  console.log(`
사용법:
  node sog-to-js.js <입력.sog> [출력.js]

설명:
  <입력.sog>  변환할 .sog 파일 경로 (필수)
  [출력.js]   생성할 .js 파일 경로 (생략 시 <입력>_sog.js)

예시:
  node sog-to-js.js assets/물문화관_3DGS_SPLAT.sog
  node sog-to-js.js assets/장미공원.sog assets/장미공원sog.js
`);
  process.exit(args.length < 1 ? 1 : 0);
}

const inputPath = path.resolve(args[0]);

// 입력 파일 존재 확인
if (!fs.existsSync(inputPath)) {
  console.error('❌ 입력 파일을 찾을 수 없습니다:', inputPath);
  process.exit(1);
}

// .sog 확장자 경고 (강제는 아님 — .ply 등 다른 자산도 같은 방식으로 내장 가능)
if (path.extname(inputPath).toLowerCase() !== '.sog') {
  console.warn('⚠️  입력 파일 확장자가 .sog 가 아닙니다:', path.basename(inputPath));
}

// 출력 경로 결정: 지정 없으면 "<입력파일명>_sog.js"
const baseName = path.basename(inputPath, path.extname(inputPath)); // 확장자 제거
const outputPath = args[1]
  ? path.resolve(args[1])
  : path.join(path.dirname(inputPath), baseName + '_sog.js');

// ---- 2) 파일 읽어 base64 인코딩 ---------------------------------------------
const buf = fs.readFileSync(inputPath);
const b64 = buf.toString('base64');

// HTML 로더(window.__SOG_NAME__)에 넘길 원본 파일 이름 (경로 제외)
const assetName = path.basename(inputPath);

// ---- 3) .js 내용 생성 --------------------------------------------------------
// 기존 자동 생성 파일과 동일한 형식 (classic script → file:// 에서도 로드됨)
const out =
  '// 이 파일은 assets/' + assetName + ' 을 base64 인코딩한 데이터입니다.\n' +
  '// (브라우저의 file:// 환경에서 직접 사용. 서버 환경에서는 원본 .sog 를 직접 불러옵니다.)\n' +
  '// 자동 생성: node sog-to-js.js — 자산 교체 시 재생성하세요.\n' +
  'window.__SOG_NAME__ = ' + JSON.stringify(assetName) + ';\n' +
  'window.__SOG_B64__ = "' + b64 + '";\n';

fs.writeFileSync(outputPath, out);

// ---- 4) 결과 보고 ------------------------------------------------------------
const mb = (n) => (n / 1024 / 1024).toFixed(2) + ' MB';
console.log('✅ 변환 완료');
console.log('   원본    :', inputPath, '(' + mb(buf.length) + ')');
console.log('   출력    :', outputPath, '(' + mb(Buffer.byteLength(out)) + ')');
console.log('   자산이름:', assetName);
console.log('');
console.log('HTML 에서 사용:');
console.log('   <script src="./assets/' + path.basename(outputPath) + '"></script>');
