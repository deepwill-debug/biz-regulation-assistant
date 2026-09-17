import os
import sys
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import List, Optional
import requests

# Windows 콘솔 UTF-8 출력 설정
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

# 사용자 OC ID (환경변수 또는 지정 ID)
OC_ID = os.environ.get("LAW_OC_ID", "incham")


@dataclass
class LawInfo:
    mst: str
    name: str
    law_type: str
    promulgation_no: str
    promulgation_date: str = ""


@dataclass
class LawArticle:
    key: str
    article_no: str
    title: str
    content: str
    sub_items: List[str]


class LawApiClient:
    """국가법령정보 공동활용 OpenAPI 클라이언트"""

    SEARCH_API_URL = "https://www.law.go.kr/DRF/lawSearch.do"
    SERVICE_API_URL = "https://www.law.go.kr/DRF/lawService.do"

    def __init__(self, oc_id: Optional[str] = None, timeout: int = 10):
        self.oc_id = oc_id or OC_ID
        self.timeout = timeout

    def search_laws(self, query: str, count: int = 5, page: int = 1) -> List[LawInfo]:
        """1단계: 키워드로 법령 목록 검색 (lawSearch.do)"""
        params = {
            "OC": self.oc_id,
            "target": "law",
            "type": "XML",
            "query": query,
            "display": count,
            "page": page,
        }

        try:
            response = requests.get(self.SEARCH_API_URL, params=params, timeout=self.timeout)
            response.raise_for_status()
            root = ET.fromstring(response.content)
        except requests.exceptions.RequestException as e:
            print(f"[오류] 법령 검색 네트워크 요청 실패: {e}")
            return []
        except ET.ParseError as e:
            print(f"[오류] 응답 XML 파싱 실패 (OC ID 또는 응답 확인 필요): {e}")
            return []

        # API 에러 응답 체크
        result_code = root.findtext(".//resultCode")
        if result_code and result_code != "00":
            result_msg = root.findtext(".//resultMsg", default="알 수 없는 오류")
            print(f"[API 오류] {result_msg} (코드: {result_code})")
            return []

        law_list = []
        for item in root.findall(".//law"):
            law_list.append(
                LawInfo(
                    mst=item.findtext("법령일련번호", default="").strip(),
                    name=item.findtext("법령명한글", default="").strip(),
                    law_type=item.findtext("법령구분명", default="").strip(),
                    promulgation_no=item.findtext("공포번호", default="").strip(),
                    promulgation_date=item.findtext("공포일자", default="").strip(),
                )
            )

        return law_list

    def get_law_articles(self, law_mst: str) -> List[LawArticle]:
        """2단계: 법령일련번호(MST)를 이용해 상세 조문 파싱 (lawService.do)"""
        params = {
            "OC": self.oc_id,
            "target": "law",
            "type": "XML",
            "MST": law_mst,
        }

        try:
            response = requests.get(self.SERVICE_API_URL, params=params, timeout=self.timeout)
            response.raise_for_status()
            root = ET.fromstring(response.content)
        except requests.exceptions.RequestException as e:
            print(f"[오류] 조문 본문 네트워크 요청 실패: {e}")
            return []
        except ET.ParseError as e:
            print(f"[오류] 조문 본문 XML 파싱 실패: {e}")
            return []

        articles = []
        for item in root.findall(".//조문단위"):
            article_key = item.findtext("조문키", default="").strip()
            article_no = item.findtext("조문번호", default="").strip()
            if article_no and not article_no.startswith("제"):
                article_no = f"제{article_no}조"

            article_title = item.findtext("조문제목", default="").strip()
            article_content = item.findtext("조문내용", default="").strip()

            sub_items = []
            # 1. 항(Paragraph) 파싱
            paragraphs = item.findall("항")
            if paragraphs:
                for p in paragraphs:
                    p_content = p.findtext("항내용", default="").strip()
                    if p_content:
                        sub_items.append(p_content)

                    # 2. 호(Item) 파싱
                    for ho in p.findall("호"):
                        ho_content = ho.findtext("호내용", default="").strip()
                        if ho_content:
                            sub_items.append(f"  {ho_content}")

                            # 3. 목(Sub-item) 파싱
                            for mok in ho.findall("목"):
                                mok_content = mok.findtext("목내용", default="").strip()
                                if mok_content:
                                    sub_items.append(f"    {mok_content}")
            else:
                # 항 없이 직접 호가 존재하는 조문 처리
                for ho in item.findall("호"):
                    ho_content = ho.findtext("호내용", default="").strip()
                    if ho_content:
                        sub_items.append(f"  {ho_content}")

            # 본문 조립 (조문내용과 항내용 간 중복 방지)
            full_text_lines = []
            if article_content:
                full_text_lines.append(article_content)

            for sub in sub_items:
                if sub not in article_content:
                    full_text_lines.append(sub)

            full_text = "\n".join(full_text_lines)

            articles.append(
                LawArticle(
                    key=article_key,
                    article_no=article_no,
                    title=article_title,
                    content=full_text,
                    sub_items=sub_items,
                )
            )

        return articles

    def search_and_extract(
        self, keyword: str, target_article_keyword: Optional[str] = None, max_laws: int = 3
    ):
        """법령 검색 및 키워드 필터링 출력"""
        print(f"\n🔍 [1단계] '{keyword}' 관련 법령 목록 검색 중...")
        laws = self.search_laws(keyword, count=max_laws)

        if not laws:
            print("❌ 검색된 법령이 없거나 API 호출에 실패했습니다. (OC 인증키 및 검색어 확인 필요)")
            return

        for law in laws:
            print(f"\n==================================================")
            print(f"📌 법령명: {law.name} ({law.law_type}) [MST: {law.mst}]")
            print(f"==================================================")

            articles = self.get_law_articles(law.mst)
            print(f"-> 총 {len(articles)}개 조문 파싱 완료.")

            matched_count = 0
            for art in articles:
                if target_article_keyword:
                    if (
                        target_article_keyword not in art.content
                        and target_article_keyword not in art.title
                    ):
                        continue

                title_str = f" {art.title}" if art.title else ""
                print(f"\n[{art.article_no}]{title_str}")
                print(art.content.strip())
                matched_count += 1

                if not target_article_keyword and matched_count >= 3:
                    print("\n... (이하 생략) ...")
                    break

            if target_article_keyword and matched_count == 0:
                print(f"※ 조문 본문 중 '{target_article_keyword}' 키워드를 포함하는 조문이 없습니다.")


if __name__ == "__main__":
    client = LawApiClient(oc_id="incham")
    client.search_and_extract(keyword="화학물질관리법", target_article_keyword="영업허가")
