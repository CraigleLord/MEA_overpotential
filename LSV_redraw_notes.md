# LSV 비교 그래프 — 작업 기록 & 재작업 가이드

최종 갱신: 2026-04-27 (matplotlib 버전 완료)
작업 폴더: `c:\Users\user\My Drive\KAIST MASc 2021\Laboratory Work\Protocol\overpotential calculation\For paper SI`

---

## 0. 빠른 재실행

```powershell
& 'C:\Users\user\AppData\Local\Programs\Python\Python312\python.exe' `
  'c:\Users\user\My Drive\KAIST MASc 2021\Laboratory Work\Protocol\overpotential calculation\For paper SI\plot_lsv_reference_style.py'
```

→ `LSV_reference_style_png/` 안에 4개 조건별 PNG 갱신.

---

## 1. 산출물 (현재 상태)

| 파일 | 설명 |
|------|------|
| [plot_lsv_reference_style.py](plot_lsv_reference_style.py) | **메인 작업 스크립트** (matplotlib). 입력: `LSV_comparison.xlsx`, 출력: PNG 4종 (조건별) + 각 패널 분리본 |
| [LSV_comparison.xlsx](LSV_comparison.xlsx) | 32개 원본 xlsx에서 추출한 (i, V) 데이터 모음. 8개 데이터 시트 (`CN_BM` ~ `AB_Polyol`) |
| [LSV_reference_style_png/](LSV_reference_style_png/) | 최종 figure 출력 폴더. 조건별 `*_combined.png` (300 dpi) + `*_LSV.png`, `*_Power.png` 분리본 |
| [build_lsv_comparison.ps1](build_lsv_comparison.ps1) | (1회용) 32개 원본 xlsx → `LSV_comparison.xlsx` 변환 스크립트. 원본 데이터 변경 시에만 재실행 필요 |
| [LSV_comparison_png/](LSV_comparison_png/) | 첫 시도 PNG 8장 (단순 4-curve overlay 스타일). 참고용으로 보관 |
| ~~build_lsv_reference_style.ps1~~ | Excel COM 시도 (axis scale 마샬링 에러로 폐기). 더 이상 사용 안 함 |

---

## 2. Figure 스펙 (스크립트에 구현된 값)

### 시리즈 (4개)
| 약어 | 샘플 | 색 (RGB 0–1) | 선 스타일 |
|------|------|--------------|-----------|
| CN BM     | BM     | (45, 160, 130)/255 — teal-green   | solid (`-`) |
| KB BM     | BM     | (120, 175, 220)/255 — powder blue | solid (`-`) |
| VC Polyol | Polyol | (0, 0, 0) — black                  | dotted, `(0,(1,1.6))` round-dot |
| AB Polyol | Polyol | (245, 180, 30)/255 — amber        | dotted, `(0,(1,1.6))` round-dot |

선 굵기: `linewidth=2.2`, `solid_capstyle="round"`, `dash_capstyle="round"`

### Max-P 마커 (LSV 좌측 패널에만)
- 위치: `(I[argmax(P)], V[argmax(P)])` — P = i × V
- `marker="o"`, `markersize=8`, facecolor/edgecolor = 시리즈 색상, `zorder=5`

### 축
| 항목 | 좌측 (LSV) | 우측 (Power) |
|------|------------|--------------|
| X 범위 | 0 → 3500 | 0 → 3500 |
| X major tick | 500 간격 | 500 간격 |
| Y 범위 | 0.2 → 1.2 | 0 → 1200 |
| Y major tick | 0.2 간격 | 200 간격 |
| X 라벨 | `Current Density (mAcm⁻²)` | `Current Density (mAcm⁻²)` |
| Y 라벨 | `Cell Voltage (V)` | `Power Density (mWcm⁻²)` |

> **참고**: 사용자가 원래 보여준 레퍼런스 이미지에서는 우측 Y축 라벨이 `Cell Voltage (V)`로 잘못 표기되어 있었음. 본 스크립트는 의미적으로 올바른 `Power Density (mWcm⁻²)`로 작성. 레퍼런스를 그대로 따르려면 [plot_lsv_reference_style.py](plot_lsv_reference_style.py)의 `axR.set_ylabel(...)` 라인을 수정.

### 스타일 디테일
- 격자선 없음 (`ax.grid(False)`)
- minor tick 끔 (`ax.minorticks_off()`)
- tick `direction="in"`, major width 1.2, length 5
- frame width: `axes.linewidth=1.4`, spine width 1.4
- 축 라벨 fontsize 13, bold; tick label 11
- 폰트: DejaVu Sans Bold (Pt 첨자 등은 mathtext)

### 범례
- 위치: 좌측 패널 좌상단 (`loc="upper left"`, `bbox_to_anchor=(0.005, 0.99)`)
- `frameon=False`, `fontsize=10.5`, `handlelength=2.4`, `labelspacing=0.35`
- 핸들은 별도 `Line2D` 리스트로 직접 구성 (max-P 마커가 범례에 안 들어가도록)

### 조건 텍스트 박스 (좌측 패널 우상단)
6줄, 우상단 정렬:
```
H₂/<Gas>          ← Air 또는 O₂ (mathtext: H$_2$/Air 또는 H$_2$/O$_2$)
RH100
Pt 5 wt%
0.05 mg_Pt/cm²    ← 0.05 mg$_{\mathrm{Pt}}$/cm$^{2}$
IC 0.8, N212
<BP> bar_g        ← f"{bar} bar$_{{g}}$" (0 또는 1.5)
```

`ax.text(0.985, 0.985, ..., transform=ax.transAxes, ha="right", va="top", fontsize=10.5, fontweight="bold", linespacing=1.35)`

### Figure 크기 / 여백
- `figsize=(11.5, 4.4)`
- `subplots_adjust(left=0.075, right=0.985, bottom=0.155, top=0.965, wspace=0.25)`
- `dpi=300`, `bbox_inches="tight"`

---

## 3. 데이터 흐름

```
원본 32개 xlsx (Overpotnital/{Air,Air BP,O2,O2 BP}/<CAT> <SAMPLE>*.xlsx)
            │
            │  (1회만) build_lsv_comparison.ps1 — Excel COM
            ▼
LSV_comparison.xlsx
  ├ CN_BM, KB_BM, VC_BM, AB_BM
  └ CN_Polyol, KB_Polyol, VC_Polyol, AB_Polyol
       각 시트 = 4개 조건 블록 (cols A:B, D:E, G:H, J:K) × (i, V)
            │
            │  plot_lsv_reference_style.py — matplotlib
            ▼
LSV_reference_style_png/<cond>_combined.png  (Air_0bp / Air_15bp / O2_0bp / O2_15bp)
                          + _LSV.png + _Power.png
```

### LSV_comparison.xlsx 시트 레이아웃
| Cols | 내용 |
|------|------|
| A:B  | Air, 0 bar BP — i (mA/cm²), V (V) |
| D:E  | Air, 1.5 bar BP — i, V |
| G:H  | O2, 0 bar BP — i, V |
| J:K  | O2, 1.5 bar BP — i, V |
- Row 1: 조건 라벨 (병합 셀)
- Row 2: 컬럼 헤더 (`i (mA/cm^2)`, `V (V)`)
- Row 3+: 데이터

### 원본 xlsx 컬럼 규칙 (Sheet1, 모든 32개 파일 공통)
- 헤더 row 1~3, 데이터 row 4~
- Column F = i (mA/cm²) ← 사용
- Column B = V (V) ← 사용
- Column A = I (A 단위), G = -V (참고용)
- 행 끝 sentinel `-2146826265` 또는 빈 셀 → 그 직전까지만 유효

---

## 4. Max Power Density 결과 (참조용)

P = i × V로 계산, argmax. 스크립트가 매번 재계산하지만 검증용으로 보관.

| 시리즈 | 조건 | P_max (mW/cm²) | i @ P_max (mA/cm²) | V @ P_max (V) |
|--------|------|----------------|--------------------|---------------|
| CN BM     | Air, 0 bar  | 263.8 | 758.5  | 0.348 |
| CN BM     | Air, 1.5 bar| 428.5 | 1089.5 | 0.393 |
| CN BM     | O2,  0 bar  | 429.3 | 1412.2 | 0.304 |
| CN BM     | O2,  1.5 bar| 647.6 | 1794.0 | 0.361 |
| KB BM     | Air, 0 bar  | 295.8 | 709.9  | 0.417 |
| KB BM     | Air, 1.5 bar| 509.6 | 1000.9 | 0.509 |
| KB BM     | O2,  0 bar  | 538.7 | 1403.9 | 0.384 |
| KB BM     | O2,  1.5 bar| 765.9 | 1704.8 | 0.449 |
| VC Polyol | Air, 0 bar  | 362.2 | 866.5  | 0.418 |
| VC Polyol | Air, 1.5 bar| 544.7 | 1237.9 | 0.440 |
| VC Polyol | O2,  0 bar  | 629.9 | 1656.9 | 0.380 |
| VC Polyol | O2,  1.5 bar| 824.0 | 2060.0 | 0.400 |
| AB Polyol | Air, 0 bar  | 403.6 | 1053.7 | 0.383 |
| AB Polyol | Air, 1.5 bar| 635.5 | 1635.1 | 0.389 |
| AB Polyol | O2,  0 bar  | 689.8 | 1839.6 | 0.375 |
| AB Polyol | O2,  1.5 bar| 967.5 | 2357.8 | 0.410 |

→ 모든 P_max < 1200, 모든 i < 3500. 현재 축 범위로 잘림 없음.

단, **O2 1.5 bar에서 AB Polyol이 i ≈ 3100 mA/cm²까지 늘어나** 우측 끝 가까이까지 도달. 더 여유 두려면 `xlim(0, 4000)`으로 변경 가능 ([plot_lsv_reference_style.py](plot_lsv_reference_style.py)의 `ax.set_xlim(0, 3500)` + `set_xticks(np.arange(0, 3501, 500))` 두 줄 수정).

---

## 5. 자주 할 만한 수정

### 시리즈 추가/교체 (예: CN을 Polyol로 바꾸기)
[plot_lsv_reference_style.py](plot_lsv_reference_style.py) 상단 `SERIES` 리스트에서 해당 튜플의 두 번째 항목(`"BM"` → `"Polyol"`) 변경.

```python
SERIES = [
    ("CN", "Polyol", (45/255, 160/255, 130/255), "-",  "CN Polyol"),  # 변경
    ...
]
```

### 색상 변경
같은 `SERIES` 튜플의 RGB 튜플 수정. 0–1 범위 정규화된 값 사용.

### 축 범위/눈금 변경
`make_figure()` 함수의 다음 블록 수정:
```python
axL.set_ylim(0.2, 1.2);  axL.set_yticks(np.arange(0.2, 1.21, 0.2))
axR.set_ylim(0,   1200); axR.set_yticks(np.arange(0,   1201, 200))
ax.set_xlim(0, 3500);    ax.set_xticks(np.arange(0, 3501, 500))
```

### Figure 크기/해상도
- `figsize=(11.5, 4.4)` → 논문 column-width 맞추려면 (5.5, 3.5) 등으로 축소
- `dpi=300` (현재 인쇄 품질). 크게 줄이고 싶으면 `dpi=150`

### 조건 박스 내용 (촉매 로딩, 멤브레인 등 변경 시)
`make_figure()`의 `annot_lines` 리스트 수정:
```python
annot_lines = [
    gas_label,
    "RH100",
    "Pt 5 wt%",
    r"0.05 mg$_{\mathrm{Pt}}$/cm$^{2}$",
    "IC 0.8, N212",
    f"{bar} bar$_{{g}}$",
]
```

### 우측 Y축 라벨을 레퍼런스 이미지처럼 "Cell Voltage (V)"로 되돌리기
```python
axR.set_ylabel("Cell Voltage (V)", fontsize=13, fontweight="bold")
```
(현재는 의미적으로 정확한 `Power Density (mWcm⁻²)`로 되어 있음)

---

## 6. 환경

- Python: `C:\Users\user\AppData\Local\Programs\Python\Python312\python.exe` (3.12.10)
- 설치된 패키지 (user-scope, `--user`): pandas 3.0.2, matplotlib 3.10.9, openpyxl 3.1.5, numpy 2.4.4
- 폰트: DejaVu Sans (matplotlib 기본). 논문 양식이 Arial/Helvetica 등을 요구하면 `plt.rcParams["font.family"]` 변경

---

## 7. 다른 사람이 이어받을 때 체크리스트

1. Python 경로가 같은지 확인 → 다르면 위 0번 명령의 경로 수정
2. `LSV_comparison.xlsx`가 있는지 확인 → 없으면 `build_lsv_comparison.ps1`로 재생성 (원본 32개 xlsx 필요)
3. 원본 xlsx의 컬럼 구조가 동일한지 (F=i, B=V) → 다른 시트면 `build_lsv_comparison.ps1` 수정 필요
4. 시리즈 4개를 다른 조합으로 바꾸려면 → §5의 "시리즈 추가/교체" 참고
