import pandas as pd

# 엑셀 파일에서 데이터 읽기
mps_df = pd.read_excel('MRP_입력정보.xlsx', sheet_name='MPS')
bom_df = pd.read_excel('MRP_입력정보.xlsx', sheet_name='BOM')
irf_df = pd.read_excel('MRP_입력정보.xlsx', sheet_name='IRF')
irf_df = irf_df.fillna(0)

# BOM 데이터를 계층 구조로 변환
def build_bom_structure(bom_df):
    bom_structure = {}
    for _, row in bom_df.iterrows():
        parent = row['Parent']
        child = row['Child']
        qty = row['Qty']
        if parent not in bom_structure:
            bom_structure[parent] = []
        bom_structure[parent].append((child, qty))
    return bom_structure

bom_structure = build_bom_structure(bom_df)

# MRP를 계산하는 함수
# 자재소요계획, 딕셔너리
def calculate_mrp(mps_df, bom_structure, irf_df):
    mrp = {}

    # 초기 입력
    for _, row in mps_df.iterrows():
        product_code = row['품목코드']
        quantity = row['수량']
        due_date = row['납기']


        # BOM 구조를 순회하며 자재 소요량 계산
        def calculate_requirements(product_code, quantity, due_date):
            if product_code in bom_structure:
                for child, qty in bom_structure[product_code]:
                    calculate_requirements(child, qty * quantity, due_date)

            if product_code not in mrp:
                mrp[product_code] = []
            mrp[product_code].append((quantity, due_date))

        # MPS 시트 초기 입력값 처리
        calculate_requirements(product_code, quantity, due_date)


    # IRF 데이터를 고려하여 주문량 계산
    for _, row in irf_df.iterrows():
        product_code = row['품목코드']
        current_inventory = row['현재재고']
        lead_time = row['인도기간']
        safety_stock = row['안전재고']
        expected_receipt_quantity = row['예정입고량']
        expected_receipt_date = row['예정입고일']
        moq = row['주문량']  # 최소주문량
        planned_order_releases = 0  # 계획발주 추가... 여기서 부터 문제
        planned_receipt = planned_order_releases # 계획입고 추가

        if product_code in mrp:
            for i in range(len(mrp[product_code])):
                quantity, due_date = mrp[product_code][i]
                if due_date > expected_receipt_date:

                    # 순소요량 계산

                    current_inventory = current_inventory - safety_stock + planned_receipt  # 추가
                    total_demand = quantity
                    expected_receipt = expected_receipt_quantity # 예정입고
                    expected_inventory = current_inventory + expected_receipt - total_demand   # 예상재고

                    net_requirements = total_demand - expected_inventory
                    if net_requirements < 0:
                        net_requirements = 0

                    if current_inventory < net_requirements:
                        planded_order_releases = max(net_requirements, moq)
                        current_inventory = current_inventory - safety_stock + planned_receipt  # 추가

                    planned_receipt = max(net_requirements, moq)
                    planned_receipt_date = expected_receipt_date
                    if lead_time == 1:
                        planned_receipt_date -= 1  # 인도기간 주단위 계획발주

                    mrp[product_code][i] = (quantity, due_date, total_demand, expected_receipt, expected_inventory, net_requirements, planned_receipt, planned_order_releases)

                elif due_date == expected_receipt_date:
                    current_inventory = current_inventory + expected_receipt_quantity + planned_receipt

    return mrp

# MRP 계산
result = calculate_mrp(mps_df, bom_structure, irf_df)

# 결과 출력

print(mps_df)
print("")
print(bom_df)
print("")
print(irf_df)
print("")
print(bom_structure)
print("")

# print("품목코드\t납기\t총소요량\t예정입고\t예상재고\t순소요량\t계획수주\t계획발주\t계획발주날짜")
print("품목코드\t납기\t총소요량\t예정입고\t예상재고\t순소요량\t계획수주\t계획발주")
for product_code, requirements in result.items():
    for requirement in requirements:
        # print(f"{product_code}\t{requirement[1]}\t{requirement[2]}\t{requirement[3]}\t{requirement[4]}\t{requirement[5]}\t{requirement[6]}\t{requirement[7]}\t{requirement[7]}")
        print(f"{product_code}\t{requirement[1]}\t{requirement[2]}\t{requirement[3]}\t{requirement[4]}\t{requirement[5]}\t{requirement[6]}\t{requirement[7]}")
