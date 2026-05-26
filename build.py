"""
정부광고 데이터 대시보드 빌드 스크립트.

원본 xlsx 8개(2019–2026)를 직접 읽어 정제·집계해 JSON으로 출력한다.

* 광고시작일 컬럼은 신뢰하지 않는다 (병합된 CSV에서 깨진 경험 + 원본도 입력 오류 다수).
  대신 **파일명에서 연도를 추출**해 행의 연도로 사용한다.
* 1차 정제: 문자열 trim, 필수 컬럼 NaN 제거, 음수/0 금액 행 제거, 광고료 + 수수료 ≠ 총계 행 플래그.

입력:
    ~/Downloads/adfile/2019~2025 ... .xlsx (7개)
    ~/Downloads/2026년도_정부광고집행내역_260430 기준.xlsx (1개)

출력:
    data/overview.json
    data/aggregate.json
    data/regime_monthly.json (매체 그룹 × 월별, 광고시작일 기준 — 정권 비교용)
    data/agency/{idx}.json (광고주별 raw 행)
    data/media/{idx}.json  (매체별 raw 행)
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pandas as pd

HOME = Path.home()
ADFILE_DIR = HOME / "Downloads" / "adfile"
EXTRA_FILES = [
    HOME / "Downloads" / "2026년도_정부광고집행내역_260430 기준.xlsx",
]
OUT_DIR = Path(__file__).resolve().parent / "data"
OUT_DIR.mkdir(parents=True, exist_ok=True)

YEAR_FROM_NAME = re.compile(r"(\d{4})")
KEY_COLS = ["기관분류", "기관명", "매체구분", "매체명"]
DETAIL_COLS_EXTRA = ["광고명", "광고료"]  # raw partition에 추가로 필요한 컬럼


def collect_files() -> list[tuple[int, Path]]:
    """파일명 첫 4자리 숫자를 연도로 보고 (year, path) 리스트 반환."""
    out: list[tuple[int, Path]] = []
    candidates = list(ADFILE_DIR.glob("*.xlsx")) + EXTRA_FILES
    for p in candidates:
        if not p.exists():
            print(f"  ⚠️  파일 없음: {p}")
            continue
        m = YEAR_FROM_NAME.search(p.name)
        if not m:
            print(f"  ⚠️  연도 추출 실패: {p.name}")
            continue
        year = int(m.group(1))
        if not (2010 <= year <= 2030):
            print(f"  ⚠️  연도 범위 밖({year}): {p.name}")
            continue
        out.append((year, p))
    out.sort()
    return out


def load_all(files: list[tuple[int, Path]]) -> pd.DataFrame:
    parts = []
    for year, path in files:
        size_mb = path.stat().st_size / 1024 / 1024
        print(f"  📂 {year} 읽는 중 ({size_mb:.1f}MB): {path.name}")
        df = pd.read_excel(path, engine="openpyxl")
        df["연도"] = year
        parts.append(df)
        print(f"     → {len(df):,}행")
    return pd.concat(parts, ignore_index=True)


def classify_regime_media(m: object) -> str | None:
    """매체명을 정권 비교용 그룹 15개로 분류. index.html의 REGIME_MEDIA_GROUPS와 동일."""
    if not isinstance(m, str):
        return None
    # 조선일보 계열
    if m.startswith("조선일보"):
        return "조선일보"
    if m.startswith(("디지틀조선일보", "월간조선", "주간조선", "스포츠조선",
                     "여성조선", "어린이조선일보", "IT조선", "조선비즈", "조선닷컴")):
        return "조선일보"
    # 중앙일보 계열
    if m.startswith("중앙일보"):
        return "중앙일보"
    if m.startswith(("월간중앙", "중앙선데이", "코리아중앙데일리", "KOREA JOONGANG DAILY")):
        return "중앙일보"
    # 동아일보 계열
    if m.startswith("동아일보"):
        return "동아일보"
    if m.startswith(("동아닷컴", "신동아", "주간동아", "스포츠동아", "IT동아",
                     "동아연감", "과학동아", "여성동아", "동아사이언스")):
        return "동아일보"
    # 한겨레 (시사한겨레/캐나다 제외)
    if m.startswith("한겨레") and "시사한겨레" not in m and "캐나다" not in m:
        return "한겨레"
    # 경향
    if m.startswith("경향신문"):
        return "경향신문"
    if m.startswith(("스포츠경향", "주간경향", "헬스경향")):
        return "경향신문"
    # 단어 경계 검사 (예: KBS, KBS-2TV는 KBS / KBSXXX는 다른 매체)
    def starts_with_word(s: str, prefix: str) -> bool:
        if not s.startswith(prefix):
            return False
        tail = s[len(prefix):]
        return tail == "" or not tail[0].isalpha()
    if starts_with_word(m, "KBS"):
        return "KBS"
    if starts_with_word(m, "MBC"):
        return "MBC"
    if starts_with_word(m, "SBS"):
        return "SBS"
    if starts_with_word(m, "EBS") or m.startswith("교육방송(EBS"):
        return "EBS"
    if starts_with_word(m, "JTBC"):
        return "JTBC"
    if m.startswith("TV조선"):
        return "TV조선"
    if m in ("채널A", "뉴스채널A") or starts_with_word(m, "채널A"):
        return "채널A"
    if starts_with_word(m, "MBN") and "세자가" not in m:
        return "MBN"
    if m.startswith("YTN"):
        return "YTN"
    if m.startswith("연합뉴스TV"):
        return "연합뉴스TV"
    return None


def build_regime_monthly(df: pd.DataFrame, out_path: Path) -> None:
    """매체 그룹 × 연월 집계. 광고시작일(YYYYMMDD 정수) 기준."""
    print("[추가] regime_monthly.json 생성 (매체 그룹 × 월)")
    sub = df.copy()
    # 광고시작일 정제 (정수형, 20000101~20301231 범위만)
    sub["광고시작일"] = pd.to_numeric(sub["광고시작일"], errors="coerce")
    sub = sub[(sub["광고시작일"] >= 20000101) & (sub["광고시작일"] <= 20301231)]
    sub["연월"] = sub["광고시작일"].astype("int64").astype(str).str[:6]

    # 매체 그룹 분류
    sub["매체그룹"] = sub["매체명"].apply(classify_regime_media)
    sub = sub[sub["매체그룹"].notna()]
    print(f"      대상 행수(매체 그룹 매칭): {len(sub):,}")

    grouped = sub.groupby(["매체그룹", "연월"])["총액"].sum().reset_index()
    print(f"      집계 (매체그룹, 연월) 조합: {len(grouped):,}")

    # 그룹별 dict로 변환: { "조선일보": { "2019-01": 12345, ... }, ... }
    result: dict[str, dict[str, int]] = {}
    for r in grouped.itertuples(index=False):
        ym_iso = f"{r.연월[:4]}-{r.연월[4:]}"  # "201901" → "2019-01"
        result.setdefault(r.매체그룹, {})[ym_iso] = int(r.총액)

    out_path.write_text(
        json.dumps(result, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )
    size_kb = out_path.stat().st_size / 1024
    print(f"      → {len(result)} 그룹, {size_kb:.1f}KB")


def clean(df: pd.DataFrame) -> pd.DataFrame:
    print(f"  원본 합계: {len(df):,}행")
    df = df.rename(columns={"총계(VAT포함)": "총액"})

    # 필수 컬럼 NaN 제거
    before = len(df)
    df = df.dropna(subset=KEY_COLS + ["총액"]).copy()
    print(f"  NaN 제거: {before - len(df):,}행 → 잔여 {len(df):,}")

    # 문자열 trim + 내부 다중공백 정리
    str_cols = KEY_COLS + [c for c in ["광고명"] if c in df.columns]
    for c in str_cols:
        df[c] = df[c].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)

    # 광고명 빈값/NaN 표기 정리 (raw 테이블에 보여줄 값)
    if "광고명" in df.columns:
        df.loc[df["광고명"].isin(["nan", "NaN", "None", ""]), "광고명"] = "(광고명 미기재)"

    # 금액 컬럼 int 변환
    df["총액"] = pd.to_numeric(df["총액"], errors="coerce").fillna(0).astype("int64")
    if "광고료" in df.columns:
        df["광고료"] = pd.to_numeric(df["광고료"], errors="coerce").fillna(0).astype("int64")

    # 0/음수 금액 제거 (집계 의미 없음)
    before = len(df)
    df = df[df["총액"] > 0]
    print(f"  0/음수 금액 제거: {before - len(df):,}행 → 잔여 {len(df):,}")

    # 빈 문자열로 강제된 키 컬럼 제거
    for c in KEY_COLS:
        df = df[df[c].str.len() > 0]
    df = df[~df["기관명"].isin(["nan", "NaN", "None"])]
    df = df[~df["매체명"].isin(["nan", "NaN", "None"])]

    print(f"  최종 정제: {len(df):,}행")
    return df


def to_python(v):
    """numpy/pandas scalar → 순수 파이썬 (json 직렬화용)."""
    if hasattr(v, "item"):
        return v.item()
    return v


def main() -> None:
    print("[1/5] 파일 수집")
    files = collect_files()
    if not files:
        raise SystemExit("xlsx 파일을 찾지 못했습니다.")
    print(f"      → {len(files)}개 파일 (연도: {[y for y, _ in files]})")

    print("[2/5] xlsx 읽기")
    df = load_all(files)

    print("[3/5] 정제")
    df = clean(df)

    print("[4/5] overview.json 생성")
    overview: dict = {}

    yearly = (
        df.groupby("연도")
        .agg(총액=("총액", "sum"), 건수=("총액", "count"))
        .reset_index()
        .sort_values("연도")
    )
    overview["yearly"] = [
        {"연도": int(r.연도), "총액": int(r.총액), "건수": int(r.건수)}
        for r in yearly.itertuples(index=False)
    ]

    mty = df.groupby(["매체구분", "연도"]).agg(총액=("총액", "sum")).reset_index()
    overview["media_type_year"] = [
        {"매체구분": r.매체구분, "연도": int(r.연도), "총액": int(r.총액)}
        for r in mty.itertuples(index=False)
    ]

    aty = df.groupby(["기관분류", "연도"]).agg(총액=("총액", "sum")).reset_index()
    overview["agency_type_year"] = [
        {"기관분류": r.기관분류, "연도": int(r.연도), "총액": int(r.총액)}
        for r in aty.itertuples(index=False)
    ]

    top_a = (
        df.groupby(["기관분류", "기관명"])
        .agg(총액=("총액", "sum"), 건수=("총액", "count"))
        .reset_index()
        .sort_values("총액", ascending=False)
        .head(30)
    )
    overview["top_agencies"] = [
        {"기관분류": r.기관분류, "기관명": r.기관명,
         "총액": int(r.총액), "건수": int(r.건수)}
        for r in top_a.itertuples(index=False)
    ]

    top_m = (
        df.groupby(["매체구분", "매체명"])
        .agg(총액=("총액", "sum"), 건수=("총액", "count"))
        .reset_index()
        .sort_values("총액", ascending=False)
        .head(30)
    )
    overview["top_media"] = [
        {"매체구분": r.매체구분, "매체명": r.매체명,
         "총액": int(r.총액), "건수": int(r.건수)}
        for r in top_m.itertuples(index=False)
    ]

    overview["kpi"] = {
        "총액": int(df["총액"].sum()),
        "총건수": int(len(df)),
        "기관수": int(df["기관명"].nunique()),
        "매체수": int(df["매체명"].nunique()),
        "평균건당": int(df["총액"].mean()),
        "기간": f"{df['연도'].min()}-{df['연도'].max()}",
    }

    (OUT_DIR / "overview.json").write_text(
        json.dumps(overview, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )

    print("[5/5] aggregate.json 생성 (검색용 인덱스 압축)")
    agg = (
        df.groupby(["연도", "기관분류", "기관명", "매체구분", "매체명"])
        .agg(총액=("총액", "sum"), 건수=("총액", "count"))
        .reset_index()
    )
    print(f"      집계 행수: {len(agg):,}")

    years = sorted(agg["연도"].unique().tolist())
    agency_types = sorted(agg["기관분류"].unique().tolist())
    media_types = sorted(agg["매체구분"].unique().tolist())
    agencies = sorted(agg["기관명"].unique().tolist())
    medias = sorted(agg["매체명"].unique().tolist())

    year_to_i = {y: i for i, y in enumerate(years)}
    at_to_i = {v: i for i, v in enumerate(agency_types)}
    mt_to_i = {v: i for i, v in enumerate(media_types)}
    a_to_i = {v: i for i, v in enumerate(agencies)}
    m_to_i = {v: i for i, v in enumerate(medias)}

    rows = [
        [
            year_to_i[r.연도],
            at_to_i[r.기관분류],
            a_to_i[r.기관명],
            mt_to_i[r.매체구분],
            m_to_i[r.매체명],
            int(r.총액),
            int(r.건수),
        ]
        for r in agg.itertuples(index=False)
    ]

    aggregate = {
        "schema": ["year", "agency_type", "agency", "media_type", "media", "amount", "count"],
        "years": years,
        "agency_types": agency_types,
        "media_types": media_types,
        "agencies": agencies,
        "media": medias,
        "rows": rows,
    }
    (OUT_DIR / "aggregate.json").write_text(
        json.dumps(aggregate, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )

    # === 정권별 비교용 월 단위 데이터 ===
    build_regime_monthly(df, OUT_DIR / "regime_monthly.json")

    # === 광고주별/매체별 raw partition ===
    print("\n[추가] 광고주별 raw JSON 생성")
    agency_dir = OUT_DIR / "agency"
    agency_dir.mkdir(exist_ok=True)
    # 기존 파일 정리 (이름 충돌/잔재 방지)
    for old in agency_dir.glob("*.json"):
        old.unlink()
    agency_to_idx = {a: i for i, a in enumerate(agencies)}
    agency_partition_count = 0
    agency_partition_bytes = 0
    for name, group in df.groupby("기관명", sort=False):
        idx = agency_to_idx.get(name)
        if idx is None:
            continue
        sub = group.sort_values("총액", ascending=False)[
            ["연도", "매체구분", "매체명", "광고명", "광고료", "총액"]
        ]
        payload = {
            "name": name,
            "type": group["기관분류"].iloc[0],
            "schema": ["year", "media_type", "media", "ad", "ad_fee", "total"],
            "rows": [
                [int(r.연도), r.매체구분, r.매체명, r.광고명, int(r.광고료), int(r.총액)]
                for r in sub.itertuples(index=False)
            ],
        }
        out = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
        (agency_dir / f"{idx}.json").write_text(out, encoding="utf-8")
        agency_partition_count += 1
        agency_partition_bytes += len(out.encode("utf-8"))
    print(f"  → {agency_partition_count:,}개 파일, 합계 {agency_partition_bytes/1024/1024:.1f}MB")

    print("\n[추가] 매체별 raw JSON 생성")
    media_dir = OUT_DIR / "media"
    media_dir.mkdir(exist_ok=True)
    for old in media_dir.glob("*.json"):
        old.unlink()
    media_to_idx = {m: i for i, m in enumerate(medias)}
    media_partition_count = 0
    media_partition_bytes = 0
    for name, group in df.groupby("매체명", sort=False):
        idx = media_to_idx.get(name)
        if idx is None:
            continue
        sub = group.sort_values("총액", ascending=False)[
            ["연도", "기관분류", "기관명", "광고명", "광고료", "총액"]
        ]
        payload = {
            "name": name,
            "type": group["매체구분"].iloc[0],
            "schema": ["year", "agency_type", "agency", "ad", "ad_fee", "total"],
            "rows": [
                [int(r.연도), r.기관분류, r.기관명, r.광고명, int(r.광고료), int(r.총액)]
                for r in sub.itertuples(index=False)
            ],
        }
        out = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
        (media_dir / f"{idx}.json").write_text(out, encoding="utf-8")
        media_partition_count += 1
        media_partition_bytes += len(out.encode("utf-8"))
    print(f"  → {media_partition_count:,}개 파일, 합계 {media_partition_bytes/1024/1024:.1f}MB")

    print("\n완료 — 파일 크기:")
    for p in sorted(OUT_DIR.glob("*.json")):
        size_mb = p.stat().st_size / 1024 / 1024
        print(f"  {p.name}: {size_mb:.2f} MB")
    print(f"  agency/: {agency_partition_count:,} files, {agency_partition_bytes/1024/1024:.1f}MB")
    print(f"  media/: {media_partition_count:,} files, {media_partition_bytes/1024/1024:.1f}MB")


if __name__ == "__main__":
    main()
