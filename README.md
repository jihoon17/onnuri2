# 상점가 위치 검색 지도 (카카오맵)

`mapData_v1.xlsx`(시장개요 / 시장별주소 / 구역)를 기반으로, 주소 또는 시장(상점가) 이름을
입력하면 해당 위치를 카카오 지도에 표시하는 정적 웹앱입니다. GitHub Pages 등 정적 호스팅에
그대로 올릴 수 있습니다.

## 파일 구성
- `index.html` — 검색창 / 검색 결과 / 지도 영역 / 지도를 볼 기준(범례) UI + 로직
- `mapData.js` — 엑셀 원본을 그대로 변환한 데이터 (시장 3곳, 필지 12건, 구역 7건)

## 1. 카카오 지도 API 키 발급
1. [카카오 개발자센터](https://developers.kakao.com) 접속 후 로그인
2. **내 애플리케이션 → 애플리케이션 추가하기**
3. 생성된 앱의 **앱 키 → JavaScript 키** 복사
4. **플랫폼 설정 → Web 플랫폼 등록**에서 실제 서비스할 도메인 등록
   - 예: `https://<GitHub 아이디>.github.io`

## 2. 키 적용
`index.html` 상단 스크립트에서 아래 줄을 찾아 발급받은 키로 교체합니다.

```js
const KAKAO_APP_KEY = "여기에_카카오_JavaScript_키를_입력하세요";
```

키를 넣지 않으면 지도 영역에 안내 문구만 표시됩니다(오류 없이 동작).

## 3. GitHub Pages 배포
1. 이 저장소(레포지토리)에 `index.html`, `mapData.js`를 루트에 커밋/푸시
2. GitHub 저장소 **Settings → Pages → Branch**에서 배포 브랜치 선택
3. 배포된 URL(`https://<아이디>.github.io/<레포명>/`)을 카카오 개발자센터 플랫폼에 등록

## 기능 요약
- **검색**: 시장(상점가)명 또는 지번 주소를 입력하면 데이터에서 일치하는 필지를 찾아 폴리곤으로 표시하고 지도 범위를 자동으로 맞춥니다.
- **데이터에 없는 주소**: 카카오 주소 검색(Geocoder)으로 대체 검색하여 위치만 이동합니다.
- **검색 결과 개수 배지**: 검색된 결과 중 `구분`이 "골목형상점가"인 시장의 개수만 뱃지로 표시합니다. (현재 데이터는 3개 시장 모두 골목형상점가입니다.)
- **지도를 볼 기준**: 필지 경계 / 구역 경계 / 둘 다 보기 전환 버튼을 우선 자리만 마련해 두었습니다. 세부 기준(줌 레벨, 표시 항목 등)은 추후 확정해서 채우면 됩니다.

## 데이터 갱신
엑셀 원본이 바뀌면 아래 스크립트로 `mapData.js`를 다시 생성할 수 있습니다.

```python
import pandas as pd, json

xl = pd.ExcelFile('mapData_v1.xlsx')
overview = pd.read_excel(xl, '시장개요')
addr = pd.read_excel(xl, '시장별주소')
zone = pd.read_excel(xl, '구역')

data = {
    'markets': [{'id': int(r['순번']), 'type': r['구분'], 'name': r['명칭']} for _, r in overview.iterrows()],
    'parcels': [{'id': int(r['순번']), 'market': r['소속시장'], 'address': r['주소'], 'coords': json.loads(r['경계좌표'])} for _, r in addr.iterrows()],
    'zones': [{'id': int(r['순번']), 'market': r['소속시장'], 'zoneNo': int(r['구역순번']), 'coords': json.loads(r['경계좌표'])} for _, r in zone.iterrows()],
}

with open('mapData.js', 'w', encoding='utf-8') as f:
    f.write('const MAP_DATA = ' + json.dumps(data, ensure_ascii=False, indent=2) + ';\n')
```
