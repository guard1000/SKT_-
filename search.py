import pandas as pd
import os, sys
import pickle
import warnings
warnings.filterwarnings("ignore")

def resource_path(relative_path):
    try:
        # PyInstaller에 의해 임시폴더에서 실행될 경우 임시폴더로 접근하는 함수
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def searcher():
    print('[1] 사번을 아래에 입력하세요. (팀 리더 또는 담당 실무자)')
    print('    입력 예시) 1112345')
    print('[2] 총 2개의 파일이 생성됩니다. (사번_Output(실본부단위).xlsx, 사번_Output(팀단위).xlsx)')
    print('    단, 직속조직의 경우 Output(팀단위) 파일만 생성됩니다.')
    print('    xlsx 파일은 현재 실행중인 .exe 파일이 위치한 동일폴더에 생성됩니다.')
    inp = input('입력: ')
    inp_사번 = int(inp)

    targets = 구성원정보_중복제거[구성원정보_중복제거['사번'] == inp_사번]  # 사번 필터
    if len(targets) == 0:  #
        print('잘못 입력하셨습니다. 사번을 확인하세요.')
    else:
        target_센터, target_실본부, target_팀 = list(targets['최상위 Lv'])[0], list(targets['실/본부 Lv'])[0], list(targets['팀 Lv'])[0]

    학습시간목표 = 100
    ##########################################
    # '[실/본부직속]', '[센터/단직속]'은 실본부가 없으므로 Output1(실본부단위) 생성 안 함
    if target_실본부 not in ['[실/본부직속]', '[센터/단직속]']:

        # 2-5. output1_대상 테이블 생성
        output1_대상 = pd.DataFrame(columns=['', ' '])
        output1_대상.loc[0] = ['사업부/센터', target_센터]
        output1_대상.loc[1] = ['실/본부', target_실본부]

        # 2-6. output1_평균학습시간 테이블 생성
        output1_평균학습시간 = pd.DataFrame(columns=['구분', '목표학습시간', '학습시간', '목표 달성률'])
        output1_평균학습시간.loc[0] = ['SK텔레콤 전사', 학습시간목표, format((전사학습시간 / 전사인원), ".3f"),
                                 (format(((전사학습시간 / 전사인원) / 학습시간목표) * 100, ".3f") + '%')]
        output1_평균학습시간.loc[1] = [target_센터, 학습시간목표,
                                 format((학습시간_센터[target_센터]['학습시간'] / 학습시간_센터[target_센터]['인원수']), ".3f"), (format(
                ((학습시간_센터[target_센터]['학습시간'] / 학습시간_센터[target_센터]['인원수']) / 학습시간목표) * 100, ".3f") + '%')]
        output1_평균학습시간.loc[2] = [target_실본부, 학습시간목표,
                                 format((학습시간_실본부[target_실본부]['학습시간'] / 학습시간_실본부[target_실본부]['인원수']), ".3f"), (format(
                ((학습시간_실본부[target_실본부]['학습시간'] / 학습시간_실본부[target_실본부]['인원수']) / 학습시간목표) * 100, ".3f") + '%')]

        # 2-7. output1_학습현황 테이블 생성
        output1_학습현황 = pd.DataFrame(columns=['순번', '실/본부', '팀', '평균학습시간', (target_실본부 + ' 평균대비')])

        특정_구성원정보 = 구성원정보[구성원정보['실/본부 Lv'] == target_실본부]  # 실본부로 필터링
        팀_names = list(set(특정_구성원정보['팀 Lv']))  # 해당 실본부에 속한 팀명 추출

        cnt_학습현황 = 0
        그룹평균 = 학습시간_실본부[target_실본부]['학습시간'] / 학습시간_실본부[target_실본부]['인원수']

        for 팀_name in 팀_names:
            팀평균 = 학습시간_팀[팀_name]['학습시간'] / 학습시간_팀[팀_name]['인원수']
            output1_학습현황.loc[cnt_학습현황] = [cnt_학습현황 + 1, target_실본부, 팀_name, format(팀평균, ".3f"),
                                          format(팀평균 - 그룹평균, ".3f")]
            cnt_학습현황 += 1

        # 2-8. output1 save file
        output1_filename = inp + '_Output(실본부단위).xlsx'
        writer = pd.ExcelWriter(output1_filename, engine='xlsxwriter')

        # 시트별 데이터 추가
        output1_대상.to_excel(writer, sheet_name='검색대상', index=False, header=False)
        output1_평균학습시간.to_excel(writer, sheet_name='평균학습시간', index=False)
        output1_학습현황.to_excel(writer, sheet_name='학습현황', index=False)

        # 시트별 옵션 지정 - 가운데 정렬 및 열 너비
        workbook = writer.book
        cell_format = workbook.add_format()
        cell_format.set_align('center')

        bold_center = workbook.add_format({'bold': True, 'align': 'center'})
        bold_center_bg = workbook.add_format({'bold': True, 'align': 'center', 'border': 2, 'bg_color': 'BDBDBD'})

        worksheet = writer.sheets['검색대상']
        worksheet.freeze_panes(1, 0)
        worksheet.set_column('A:A', 15, bold_center)
        worksheet.set_column('B:B', 30, cell_format)

        worksheet.write('A1', '사업부/센터', bold_center_bg)
        worksheet.write('A2', '실/본부', bold_center_bg)

        worksheet = writer.sheets['평균학습시간']
        worksheet.freeze_panes(1, 0)
        worksheet.set_column('A:A', 30, cell_format)
        worksheet.set_column('B:B', 15, cell_format)
        worksheet.set_column('C:C', 15, cell_format)
        worksheet.set_column('D:D', 15, cell_format)

        worksheet.write('A1', '구분', bold_center_bg)
        worksheet.write('B1', '목표학습시간', bold_center_bg)
        worksheet.write('C1', '학습시간', bold_center_bg)
        worksheet.write('D1', '목표 달성률', bold_center_bg)

        worksheet = writer.sheets['학습현황']
        worksheet.freeze_panes(1, 0)
        worksheet.set_column('A:A', 5, cell_format)
        worksheet.set_column('B:B', 30, cell_format)
        worksheet.set_column('C:C', 30, cell_format)
        worksheet.set_column('D:D', 15, cell_format)
        worksheet.set_column('E:E', 30, cell_format)

        worksheet.write('A1', '순번', bold_center_bg)
        worksheet.write('B1', '실/본부', bold_center_bg)
        worksheet.write('C1', '팀', bold_center_bg)
        worksheet.write('D1', '평균학습시간', bold_center_bg)
        worksheet.write('E1', target_실본부 + ' 평균대비', bold_center_bg)

        # 엑셀 파일 저장
        writer.save()
        print(output1_filename, '  파일 생성 완료.')

    # 2-9. output2_대상 테이블 생성
    output2_대상 = pd.DataFrame(columns=['', ' '])
    output2_대상.loc[0] = ['사업부/센터', target_센터]
    output2_대상.loc[1] = ['실/본부', target_실본부]
    output2_대상.loc[2] = ['팀', target_팀]
    output2_대상.loc[3] = ['팀인원', 학습시간_팀[target_팀]['인원수']]

    # 2-10. output2_평균학습시간 테이블 생성
    output2_평균학습시간 = pd.DataFrame(columns=['구분', '목표학습시간', '학습시간', '목표 달성률'])
    output2_평균학습시간.loc[0] = ['SK텔레콤 전사', 학습시간목표, format((전사학습시간 / 전사인원), ".3f"),
                             (format(((전사학습시간 / 전사인원) / 학습시간목표) * 100, ".3f") + '%')]
    output2_평균학습시간.loc[1] = [target_센터, 학습시간목표, format((학습시간_센터[target_센터]['학습시간'] / 학습시간_센터[target_센터]['인원수']), ".3f"),
                             (format(((학습시간_센터[target_센터]['학습시간'] / 학습시간_센터[target_센터]['인원수']) / 학습시간목표) * 100,
                                     ".3f") + '%')]
    if target_실본부 not in ['[실/본부직속]', '[센터/단직속]']:
        output2_평균학습시간.loc[2] = [target_실본부, 학습시간목표,
                                 format((학습시간_실본부[target_실본부]['학습시간'] / 학습시간_실본부[target_실본부]['인원수']), ".3f"), (format(
                ((학습시간_실본부[target_실본부]['학습시간'] / 학습시간_실본부[target_실본부]['인원수']) / 학습시간목표) * 100, ".3f") + '%')]
    else:
        output2_평균학습시간.loc[2] = [target_실본부, '-', '-', '-']
    output2_평균학습시간.loc[3] = [target_팀, 학습시간목표, format((학습시간_팀[target_팀]['학습시간'] / 학습시간_팀[target_팀]['인원수']), ".3f"), (
                format(((학습시간_팀[target_팀]['학습시간'] / 학습시간_팀[target_팀]['인원수']) / 학습시간목표) * 100, ".3f") + '%')]

    # 2-11. output2_학습현황 테이블 생성
    output2_학습현황 = pd.DataFrame(columns=['순번', '사번', '성명', '총 학습시간', 'mySUNI 학습시간', 'TLP 학습시간'])

    특정_구성원정보 = 구성원정보_중복제거[구성원정보_중복제거['팀 Lv'] == target_팀]  # 팀으로 필터링
    특정_구성원정보 = 특정_구성원정보.sort_values('사번')  # 사번순 정렬

    for _ in range(len(특정_구성원정보)):
        순번, 사번, 성명 = _ + 1, 특정_구성원정보.iloc[_]['사번'], 특정_구성원정보.iloc[_]['성명']
        총학습, mySUNI, TLP = 0, 0, 0

        output_특정인 = output_종합[output_종합['사번'] == str(특정_구성원정보.iloc[_]['사번'])]  # 사번으로 output_종합 필터링
        if len(output_특정인) != 0:
            총학습 = sum(output_특정인['학습시간(시간)'])

            output_특정인_mySUNI = output_특정인[output_특정인['구분'] == 'mySUNI']
            if len(output_특정인_mySUNI) != 0:
                mySUNI = sum(output_특정인_mySUNI['학습시간(시간)'])

            output_특정인_TLP = output_특정인[output_특정인['구분'] == 'TLP']
            if len(output_특정인_TLP) != 0:
                TLP = sum(output_특정인_TLP['학습시간(시간)'])

        output2_학습현황.loc[_] = [순번, 사번, 성명, format(총학습, ".3f"), format(mySUNI, ".3f"), format(TLP, ".3f")]

    # 2-12. output2_학습이력 테이블 생성
    output2_학습이력 = pd.DataFrame(columns=['순번', '사번', '이름', '과정명', '학습시간', '학습구분'])

    output_이력 = output_종합[output_종합['팀 Lv'] == target_팀]
    output_이력['사번'] = [int(_) for _ in list(output_이력['사번'])]
    output_이력 = output_이력.sort_values('사번')  # 사번순 정렬

    for _ in range(len(output_이력)):
        output2_학습이력.loc[_] = [_ + 1, output_이력.iloc[_]['사번'], output_이력.iloc[_]['성명'], output_이력.iloc[_]['과정명'],
                               output_이력.iloc[_]['학습시간(시간)'], output_이력.iloc[_]['구분']]

    # 2-13. output2 save file
    output2_filename = inp + '_Output(팀단위).xlsx'
    writer = pd.ExcelWriter(output2_filename, engine='xlsxwriter')

    # 시트별 데이터 추가
    output2_대상.to_excel(writer, sheet_name='검색대상', index=False, header=False)
    output2_평균학습시간.to_excel(writer, sheet_name='평균학습시간', index=False)
    output2_학습현황.to_excel(writer, sheet_name='학습현황', index=False)
    output2_학습이력.to_excel(writer, sheet_name='학습이력', index=False)

    # 시트별 옵션 지정 - 가운데 정렬 및 열 너비
    workbook = writer.book
    cell_format = workbook.add_format()
    cell_format.set_align('center')
    bold_center = workbook.add_format({'bold': True, 'align': 'center'})
    bold_center_bg = workbook.add_format({'bold': True, 'align': 'center', 'border': 2, 'bg_color': 'BDBDBD'})

    worksheet = writer.sheets['검색대상']
    worksheet.freeze_panes(1, 0)
    worksheet.set_column('A:A', 15, bold_center)
    worksheet.set_column('B:B', 30, cell_format)

    worksheet.write('A1', '사업부/센터', bold_center_bg)
    worksheet.write('A2', '실/본부', bold_center_bg)
    worksheet.write('A3', '팀', bold_center_bg)
    worksheet.write('A4', '팀인원', bold_center_bg)

    worksheet = writer.sheets['평균학습시간']
    worksheet.freeze_panes(1, 0)
    worksheet.set_column('A:A', 30, cell_format)
    worksheet.set_column('B:B', 15, cell_format)
    worksheet.set_column('C:C', 15, cell_format)
    worksheet.set_column('D:D', 15, cell_format)

    worksheet.write('A1', '구분', bold_center_bg)
    worksheet.write('B1', '목표학습시간', bold_center_bg)
    worksheet.write('C1', '학습시간', bold_center_bg)
    worksheet.write('D1', '목표 달성률', bold_center_bg)

    worksheet = writer.sheets['학습현황']
    worksheet.freeze_panes(1, 0)
    worksheet.set_column('A:A', 5, cell_format)
    worksheet.set_column('B:B', 10, cell_format)
    worksheet.set_column('C:C', 15, cell_format)
    worksheet.set_column('D:D', 15, cell_format)
    worksheet.set_column('E:E', 15, cell_format)
    worksheet.set_column('F:F', 15, cell_format)

    worksheet.write('A1', '순번', bold_center_bg)
    worksheet.write('B1', '사번', bold_center_bg)
    worksheet.write('C1', '성명', bold_center_bg)
    worksheet.write('D1', '총 학습시간', bold_center_bg)
    worksheet.write('E1', 'mySUNI 학습시간', bold_center_bg)
    worksheet.write('F1', 'TLP 학습시간', bold_center_bg)

    worksheet = writer.sheets['학습이력']
    worksheet.freeze_panes(1, 0)
    worksheet.set_column('A:A', 5, cell_format)
    worksheet.set_column('B:B', 10, cell_format)
    worksheet.set_column('C:C', 15, cell_format)
    worksheet.set_column('D:D', 50)
    worksheet.set_column('E:E', 15, cell_format)
    worksheet.set_column('F:F', 15, cell_format)

    worksheet.write('A1', '순번', bold_center_bg)
    worksheet.write('B1', '사번', bold_center_bg)
    worksheet.write('C1', '이름', bold_center_bg)
    worksheet.write('D1', '과정명', bold_center_bg)
    worksheet.write('E1', '학습시간', bold_center_bg)
    worksheet.write('F1', '학습구분', bold_center_bg)

    # 엑셀 파일 저장
    writer.save()
    print(output2_filename, ' 파일 생성 완료.')

##################################################################

if __name__ == "__main__":
    print('================================================')
    print('[SKT 학습시간 현황 파악]을 위한 검색기입니다.')
    print('                최종 수정 일자 : 2021.04.26')
    print('================================================\n\n')

    # laod data
    학습시간_path = resource_path('sources')
    with open(학습시간_path, 'rb') as f:
        data = pickle.load(f)
    전사학습시간 = data['전사학습시간']
    전사인원 = data['전사인원']

    학습시간_센터 = data['학습시간_센터']
    학습시간_실본부 = data['학습시간_실본부']
    학습시간_팀 = data['학습시간_팀']

    종합_path = resource_path('output_종합.csv')
    output_종합 = pd.read_csv(종합_path)
    output_종합 = output_종합[output_종합['팀 Lv'] != '구성원정보_미등록자']  # 구성원정보 미등록자는 제외
    output_종합['사번'] = [str(_) for _ in list(output_종합['사번'])]

    구성원_path = resource_path('구성원정보.csv')
    구성원정보 = pd.read_csv(구성원_path)
    구성원정보_중복제거 = 구성원정보.drop_duplicates(['사번'], keep='first')

    while True:
        searcher()
        print()


