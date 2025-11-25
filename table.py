import pandas as pd

# 엑셀 파일 이름 입력 받기
file_name = input("엑셀 파일 이름을 입력하세요 (예: data.xlsx): ")

try:
    df = pd.read_excel(file_name)
except Exception as e:
    print("파일을 읽는 중 오류가 발생했습니다:", e)
    print("data.xlsx 파일을 다시 시도합니다.")
    df = pd.read_excel("data.xlsx")

print("=== 원본 데이터(앞 5행) ===")
print(df.head())
print("컬럼 목록:", df.columns.tolist())
print()

df = df.fillna(0)

# 컬럼 이름 공백 제거
df.columns = [str(c).strip() for c in df.columns]

print("=== 컬럼 이름 정리 후 목록 ===")
print(df.columns.tolist())
print()

# 항목 / 지수 / 값 
item_col = None
index_col = None
value_col = None

for col in df.columns:
    col_str = str(col)
    low = col_str.lower()
    # 항목 후보
    if item_col is None:
        if ("항목" in col_str) or ("품목" in col_str) or ("분류" in col_str):
            item_col = col
    # 지수 후보
    if index_col is None:
        if ("지수" in col_str) or ("코드" in col_str) or ("index" in low):
            index_col = col
    # 값 후보
    if value_col is None:
        if ("값" in col_str) or ("value" in low) or ("수치" in col_str):
            value_col = col


if item_col is None and len(df.columns) > 0:
    item_col = df.columns[0]
if index_col is None and len(df.columns) > 1:
    index_col = df.columns[1]
if value_col is None:
    if len(df.columns) > 2:
        value_col = df.columns[2]
    else:
        value_col = df.columns[-1]

print("=== 사용될 컬럼 이름 ===")
print("항목 컬럼:", item_col)
print("지수 컬럼:", index_col)
print("값   컬럼:", value_col)
print()

# 타입 정리 (항목, 지수는 문자열 / 값은 숫자)
try:
    df[item_col] = df[item_col].astype(str)
except Exception as e:
    print("항목 컬럼 형변환 중 오류:", e)

try:
    df[index_col] = df[index_col].astype(str)
except Exception as e:
    print("지수 컬럼 형변환 중 오류:", e)

df[value_col] = pd.to_numeric(df[value_col], errors="coerce")
df[value_col] = df[value_col].fillna(0)

# 값 컬럼의 음수 데이터 존재 여부 체크
try:
    negative_count = (df[value_col] < 0).sum()
    if negative_count > 0:
        print("=== 경고: 음수 값 존재 ===")
        print("값 컬럼에 음수 데이터 개수:", int(negative_count))
        print("데이터 혹은 집계 로직을 한 번 더 확인해 보세요.")
        print()
except Exception as e:
    print("음수 값 체크 중 오류:", e)
    print()

print("=== 기본 정리 후 데이터(앞 5행) ===")
try:
    print(df[[item_col, index_col, value_col]].head())
except Exception:
    print(df.head())
print()

# 항목/지수별로 값 집계 (평균 사용)
try:
    grouped = df.groupby([item_col, index_col])[value_col].mean().reset_index()
except Exception as e:
    print("그룹 집계 중 오류:", e)
    grouped = df.copy()

print("=== 항목/지수별 집계 결과(일부) ===")
print(grouped.head(20))
print("집계 행 개수:", len(grouped))
print()

# 항목별 요약 통계(건수, 평균, 최소, 최대) 계산
try:
    summary = grouped.groupby(item_col)[value_col].agg(["count", "mean", "min", "max"]).reset_index()
    print("=== 항목별 요약 통계(일부) ===")
    print(summary.head(20))
    print()
except Exception as e:
    print("항목별 요약 통계 계산 중 오류:", e)
    summary = pd.DataFrame()

# 항목별 전체
try:
    item_total = grouped.groupby(item_col)[value_col].sum().reset_index()
    item_total[index_col] = "총합"   
    # 컬럼 순서 맞추기
    item_total = item_total[[item_col, index_col, value_col]]
except Exception as e:
    print("항목별 합계 계산 중 오류:", e)
    item_total = pd.DataFrame(columns=[item_col, index_col, value_col])

# 값이 큰 상위 N개 항목 출력 (총합 기준)
TOP_N = 5
try:
    # item_total에는 index_col이 "총합"으로 들어가 있으므로 그대로 사용
    top_items = item_total.sort_values(value_col, ascending=False).head(TOP_N)
    print(f"=== 값이 큰 상위 {TOP_N}개 항목(총합 기준) ===")
    print(top_items[[item_col, value_col]])
    print()
except Exception as e:
    print("상위 항목 계산 중 오류:", e)
    print()

# 상세 데이터 합치기
try:
    detail = pd.concat([grouped, item_total], ignore_index=True)
except Exception as e:
    print("데이터 병합 중 오류:", e)
    detail = grouped

# 정렬
try:
    detail = detail.sort_values([item_col, index_col])
except Exception:
    pass

print("=== 합쳐진 상세 데이터(앞 30행) ===")
print(detail.head(30))
print("상세 행 개수:", len(detail))
print()

# 피벗 테이블 만들기 (행: 항목, 열: 지수, 값: value_col)
try:
    pivot = detail.pivot(index=item_col, columns=index_col, values=value_col)
    pivot = pivot.fillna(0)
except Exception as e:
    print("피벗 테이블 생성 중 오류가 발생했습니다:", e)
    pivot = detail

print("=== 피벗 테이블(앞 부분) ===")
try:
    print(pivot.head())
except Exception:
    print(pivot)
print()

# 피벗 테이블에 행/열 합계 추가
try:
    pivot_with_total = pivot.copy()
    # 행 기준 합계
    pivot_with_total["행합계"] = pivot_with_total.sum(axis=1)
    # 열 기준 합계(마지막에 합계 행 추가)
    total_row = pivot_with_total.sum(axis=0)
    total_row.name = "열합계"
    pivot_with_total = pd.concat([pivot_with_total, total_row.to_frame().T])
    print("=== 합계가 포함된 피벗 테이블(앞 부분) ===")
    print(pivot_with_total.head())
    print()
except Exception as e:
    print("피벗 테이블 합계 추가 중 오류:", e)
    pivot_with_total = pivot