import zipfile
from lxml import etree

# XML名前空間の設定
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def get_paragraph_text_lxml(p, mode='after'):
    """
    mode:
      'before' : 変更前のテキスト（w:ins内の追加部分は除外、w:del内の削除部分は含む）
      'after'  : 変更後のテキスト（w:del内の削除部分は除外、w:ins内の追加部分は含む）
    """
    texts = []
    for elem in p.iter():
        if elem.tag.endswith('}t'):  # テキスト要素
            # 祖先にw:ins, w:delがあるかチェック
            in_ins = any(ancestor.tag.endswith('}ins') for ancestor in elem.iterancestors())
            in_del = any(ancestor.tag.endswith('}del') for ancestor in elem.iterancestors())
            if mode == 'before':
                # 変更前：挿入されたテキストは含めない
                if not in_ins:
                    texts.append(elem.text or '')
            else:  # mode == 'after'
                # 変更後：削除されたテキストは含めない
                if not in_del:
                    texts.append(elem.text or '')
    return ''.join(texts)

def extract_changed_paragraphs(docx_path):
    """
    指定したdocxファイルから、変更（挿入 or 削除）のある段落だけを抽出し、
    各段落について変更前と変更後のテキストを返す。
    戻り値は (変更前テキスト, 変更後テキスト) のタプルのリスト。
    """
    # ZIPとしてdocxを開く
    with zipfile.ZipFile(docx_path) as z:
        xml_content = z.read('word/document.xml')
    
    # XMLパース
    parser = etree.XMLParser(ns_clean=True)
    tree = etree.fromstring(xml_content, parser)
    
    # 全ての段落を取得
    paragraphs = tree.xpath('//w:p', namespaces=NS)
    
    
    results = []
    for p in paragraphs:
        # 段落内に<w:ins>または<w:del>が存在する場合、変更があったと判断
        has_ins = p.xpath('.//w:ins', namespaces=NS)
        has_del = p.xpath('.//w:del', namespaces=NS)
        print(has_ins)
        print(has_del)
        if has_ins or has_del:
            before_text = get_paragraph_text_lxml(p, mode='before')
            after_text = get_paragraph_text_lxml(p, mode='after')
            results.append((before_text, after_text))
    return results

if __name__ == '__main__':
    # 処理対象のdocxファイルのパス
    docx_path = '000087509.docx'
    changed_paragraphs = extract_changed_paragraphs(docx_path)
    
    # 各段落ごとに変更前と変更後のテキストを出力
    for idx, (before, after) in enumerate(changed_paragraphs, start=1):
        print(f"段落 {idx}")
        print("【変更前】")
        print(before)
        print("【変更後】")
        print(after)
        print("-------------------------------")
        