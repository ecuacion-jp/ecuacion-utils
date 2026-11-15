#
# Copyright © 2012 ecuacion.jp (info@ecuacion.jp)
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#

import zipfile, sys
W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
def run(t, rpr=''):
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % ("<w:rPr>"+rpr+"</w:rPr>" if rpr else "", t)
def p(inner, ppr=''):
    return '<w:p>%s%s</w:p>' % ("<w:pPr>"+ppr+"</w:pPr>" if ppr else "", inner)
def fld(instr, cached):
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> %s </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>%s</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>') % (instr, cached)
articles=[('目的','本契約は、甲が乙に対して委託する業務（以下「本業務」という。）の内容、条件その他必要な事項を定めることを目的とする。'),
 ('定義','本契約において使用する用語の定義は、次の各号に定めるとおりとする。'),
 ('業務の内容','乙は、善良な管理者の注意をもって本業務を遂行するものとし、甲の事前の書面による承諾なく、本業務の全部又は一部を第三者に再委託してはならない。'),
 ('委託料及び支払方法','甲は、乙に対し、本業務の対価として別表に定める委託料を、毎月末日締め翌月末日までに、乙の指定する銀行口座に振り込む方法により支払う。振込手数料は甲の負担とする。'),
 ('秘密保持','甲及び乙は、本契約の履行に関連して知り得た相手方の技術上又は営業上の情報（以下「秘密情報」という。）を、相手方の事前の書面による承諾なく第三者に開示又は漏洩してはならず、本契約の目的以外に使用してはならない。'),
 ('知的財産権','本業務の成果物に関する著作権（著作権法第27条及び第28条の権利を含む。）その他一切の知的財産権は、委託料の完済をもって乙から甲に移転する。'),
 ('損害賠償','甲又は乙は、本契約に違反して相手方に損害を与えた場合、相手方に対し、当該損害（合理的な弁護士費用を含む。）を賠償する責任を負う。ただし、賠償額は、当該損害の直接の原因となった個別契約の委託料総額（税込 1,000,000 円を上限とする。）を超えないものとする。'),
 ('契約期間','本契約の有効期間は、2026年4月1日から2027年3月31日までとする。ただし、期間満了の3か月前までに甲乙いずれからも書面による別段の意思表示がないときは、本契約は同一条件でさらに1年間更新されるものとし、以後も同様とする。'),
 ('解除','甲又は乙は、相手方が次の各号のいずれかに該当したときは、何らの催告を要せず、直ちに本契約の全部又は一部を解除することができる。'),
 ('反社会的勢力の排除','甲及び乙は、自己又は自己の役員が、暴力団、暴力団員、暴力団関係企業、総会屋その他これらに準ずる者（以下「反社会的勢力」という。）に該当しないこと、及び将来にわたっても該当しないことを表明し、保証する。'),
 ('協議事項','本契約に定めのない事項又は本契約の解釈に疑義が生じた事項については、甲乙誠意をもって協議の上、これを解決する。'),
 ('合意管轄','本契約に関する一切の紛争については、東京地方裁判所を第一審の専属的合意管轄裁判所とする。')]
items=['「業務委託料」とは、Service Level Agreement（SLA）に基づき算定される金額をいう。','「営業日」とは、土曜日、日曜日、国民の祝日及び12月29日から1月3日までを除く日をいう。','「成果物」とは、乙が本業務の遂行の過程で作成し、甲に納入する一切の資料、プログラム及びデータをいう。']
body=[]
body.append(p(run('業務委託基本契約書'),'<w:pStyle w:val="Title"/>'))
body.append(p('', ''))
toc=[]
for i,(t,_) in enumerate(articles):
    first = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-1" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' if i==0 else ''
    toc.append(p(first+'<w:hyperlink w:anchor="_Toc%d" w:history="1">%s<w:r><w:tab/></w:r>%s</w:hyperlink>' % (i, run("第%d条　%s"%(i+1,t)), fld("PAGEREF _Toc%d \\h"%i,"99")),
        '<w:pStyle w:val="TOC1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8494"/></w:tabs>'))
toc.append(p('<w:r><w:fldChar w:fldCharType="end"/></w:r>'))
body.append('<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr><w:sdtContent>'+p(run('目次'),'<w:pStyle w:val="TOCHeading"/>')+''.join(toc)+'</w:sdtContent></w:sdt>')
body.append(p('<w:r><w:br w:type="page"/></w:r>'))
body.append(p(run('株式会社サンプル商事（以下「甲」という。）と、株式会社エクアシオン（以下「乙」という。）とは、甲が乙に委託する業務に関し、次のとおり業務委託基本契約（以下「本契約」という。）を締結する。なお、本契約は ')+run('Master Service Agreement','<w:i/>')+run(' に相当するものであり、')+run('個別契約','<w:b/>')+run('に優先して適用される')+run('※1','<w:vertAlign w:val="superscript"/>')+run('。'),
    '<w:ind w:firstLineChars="100" w:firstLine="210"/>'))
def tc(t,w,extra='',rpr='',ppr=''):
    return '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s</w:tcPr>%s</w:tc>' % (w, extra, p(run(t,rpr),ppr))
for rep in range(2):
  for i,(t,txt) in enumerate(articles):
    if rep==1 and i>5: break
    bm = '<w:bookmarkStart w:id="%d" w:name="_Toc%d"/>' % (i,i) if rep==0 else ''
    bme = '<w:bookmarkEnd w:id="%d"/>' % i if rep==0 else ''
    body.append(p(bm+run(t)+bme,'<w:pStyle w:val="1"/>'))
    body.append(p(run(txt),'<w:pStyle w:val="2"/>'))
    if i in (1,8):
        for it in items:
            body.append(p(run(it),'<w:pStyle w:val="3"/>'))
    else:
        body.append(p(run('前項の規定にかかわらず、甲及び乙は、書面（電磁的方法を含む。）による合意により、本条の定めを変更することができる。'),'<w:pStyle w:val="2"/>'))
    if i==3:
        shd='<w:shd w:val="clear" w:fill="D9E2F3"/>'
        right='<w:jc w:val="right"/>'
        rows=[]
        rows.append('<w:tr><w:trPr><w:tblHeader/></w:trPr>'+tc('区分',1700,shd,'<w:b/>')+tc('内容',4200,shd,'<w:b/>')+tc('月額（税抜）',2594,shd,'<w:b/>')+'</w:tr>')
        rows.append('<w:tr>'+tc('基本業務',1700,'<w:vMerge w:val="restart"/><w:vAlign w:val="center"/>')+tc('システム運用保守（平日 9:00〜18:00）',4200)+tc('500,000 円',2594,'','',right)+'</w:tr>')
        rows.append('<w:tr>'+tc('',1700,'<w:vMerge/>')+tc('問い合わせ対応（月 20 件まで。超過分は別途見積りとし、甲乙協議の上で決定する。）',4200)+tc('120,000 円',2594,'','',right)+'</w:tr>')
        rows.append('<w:tr>'+tc('合計',5900,'<w:gridSpan w:val="2"/>','<w:b/>')+tc('620,000 円',2594,'','<w:b/>',right)+'</w:tr>')
        body.append('<w:tbl><w:tblPr><w:tblStyle w:val="a3"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="1700"/><w:gridCol w:w="4200"/><w:gridCol w:w="2594"/></w:tblGrid>'+''.join(rows)+'</w:tbl>')
        body.append(p(''))
body.append(p(run('本契約の成立を証するため、本書2通を作成し、甲乙記名押印の上、各1通を保有する。'),'<w:spacing w:beforeLines="100" w:before="360"/><w:ind w:firstLineChars="100" w:firstLine="210"/>'))
body.append(p(run('2026年4月1日'),'<w:jc w:val="right"/>'))
sigtabs='<w:tabs><w:tab w:val="left" w:pos="4200"/><w:tab w:val="left" w:pos="7900"/></w:tabs>'
sigind='<w:ind w:left="3360"/>'
for who in ('甲','乙'):
    body.append(p(run(who)+'<w:r><w:tab/></w:r>'+run('住所　')+'<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:tab/></w:r>', sigtabs+'<w:spacing w:beforeLines="50" w:before="180"/>'+sigind))
    body.append(p('<w:r><w:tab/></w:r>'+run('氏名　')+'<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:tab/></w:r>'+run('　印'), sigtabs+sigind))
sect='<w:sectPr><w:headerReference w:type="default" r:id="rId10"/><w:footerReference w:type="default" r:id="rId11"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701" w:header="851" w:footer="992" w:gutter="0"/><w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>'
document='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s%s</w:body></w:document>' % (W, "".join(body), sect)
FONT=sys.argv[2] if len(sys.argv)>2 else '游明朝'
styles='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Century" w:eastAsia="%s" w:hAnsi="Century"/><w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="a"/><w:pPr><w:spacing w:before="240" w:after="120"/><w:jc w:val="center"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="游ゴシック" w:hAnsi="Arial"/><w:b/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/><w:next w:val="a"/><w:pPr><w:keepNext/><w:numPr><w:numId w:val="1"/></w:numPr><w:spacing w:beforeLines="50" w:before="180"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="游ゴシック" w:hAnsi="Arial"/><w:b/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 2"/><w:basedOn w:val="a"/><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="3"><w:name w:val="heading 3"/><w:basedOn w:val="a"/><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="a"/></w:style>
<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="a"/><w:pPr><w:spacing w:after="120"/><w:jc w:val="left"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="a"/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4252"/><w:tab w:val="right" w:pos="8504"/></w:tabs><w:snapToGrid w:val="0"/></w:pPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="table" w:default="1" w:styleId="a1"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="a3"><w:name w:val="Table Grid"/><w:basedOn w:val="a1"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>''' % (W, FONT)
numbering='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering %s>
<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="1"/><w:suff w:val="space"/><w:lvlText w:val="第%%1条"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimalFullWidth"/><w:pStyle w:val="2"/><w:lvlText w:val="%%2"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="420" w:hanging="420"/></w:pPr></w:lvl>
<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="3"/><w:lvlText w:val="(%%3)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="840" w:hanging="420"/></w:pPr></w:lvl>
</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>''' % W
settings='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:defaultTabStop w:val="840"/><w:characterSpacingControl w:val="compressPunctuation"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/></w:settings>' % W
theme='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック"/></a:majorFont><a:minorFont><a:latin typeface="Century"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="%s"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>' % FONT
header='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr %s>%s</w:hdr>' % (W, p(run("業務委託基本契約書",'<w:color w:val="808080"/><w:sz w:val="18"/>'),'<w:pStyle w:val="Header"/><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="808080"/></w:pBdr><w:jc w:val="right"/>'))
footer='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr %s>%s</w:ftr>' % (W, p(fld("PAGE","1")+run(" / ")+fld("NUMPAGES","1"),'<w:jc w:val="center"/>'))
ct='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'''
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="%s/officeDocument" Target="word/document.xml"/></Relationships>' % R
drels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="{R}/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="{R}/settings" Target="settings.xml"/><Relationship Id="rId10" Type="{R}/header" Target="header1.xml"/><Relationship Id="rId11" Type="{R}/footer" Target="footer1.xml"/></Relationships>'.replace('{R}', R)
with zipfile.ZipFile(sys.argv[1],'w',zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml',ct); z.writestr('_rels/.rels',rels); z.writestr('word/_rels/document.xml.rels',drels)
    z.writestr('word/document.xml',document); z.writestr('word/styles.xml',styles); z.writestr('word/numbering.xml',numbering)
    z.writestr('word/settings.xml',settings);  z.writestr('word/header1.xml',header); z.writestr('word/footer1.xml',footer)
