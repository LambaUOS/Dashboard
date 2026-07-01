const fs = require('fs');
const path = require('path');

const srcName = '물문화관';
const srcPath = path.join(__dirname, 'assets', srcName);
const outPath = path.join(__dirname, 'assets', 'sog-data.js');

const buf = fs.readFileSync(srcPath);
const b64 = buf.toString('base64');

const out =
  '// 자동 생성 파일 — assets/' + srcName + ' 의 base64 내장 데이터\n' +
  '// (오프라인 file:// 더블클릭 지원용. 자산 교체 시 gen-sog-data.js 로 재생성)\n' +
  'window.__SOG_NAME__ = ' + JSON.stringify(srcName) + ';\n' +
  'window.__SOG_B64__ = "' + b64 + '";\n';

fs.writeFileSync(outPath, out);
console.log('원본:', (buf.length / 1024 / 1024).toFixed(1), 'MB ->',
            'sog-data.js:', (out.length / 1024 / 1024).toFixed(1), 'MB');
