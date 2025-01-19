from difflib import SequenceMatcher
# 문자열 유사도를 계산하는 함수
def similarity_score(str1, str2):
    return SequenceMatcher(None, str1, str2).ratio()

# 매칭 알고리즘
def match_wine_names(list_a, list_b):
    result = {}
    list_b_used = set()  # 이미 매칭된 항목 추적

    for name_a in list_a:
        best_match = None
        best_score = 0

        for name_b in list_b:
            if name_b in list_b_used:
                continue  # 이미 매칭된 항목은 건너뜀
            
            # 문자열 유사도 계산
            score = similarity_score(name_a.replace(' ', ''), name_b.replace(' ', ''))
            
            # 가장 높은 유사도를 가진 항목 찾기
            if score > best_score:
                best_score = score
                best_match = name_b

        if best_match:
            result[name_a] = best_match
            list_b_used.add(best_match)  # 매칭된 항목은 추적

    return result
