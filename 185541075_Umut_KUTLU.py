from docx import *
import docx.package
import docx.parts.document
import docx.parts.numbering
import logging
import json

#İçerisine verdiğimiz style keywordleri ile style nesnelerini alıyoruz
def getNecessarryStyles(styles,keywords):
    x = []
    for s in styles:
        c = str(s)
        for k in keywords:
            if k in c:
                x.append(s)
                break
    return x
#Kaynakları almak için önce kaynaklar başlığının olduğu paragrafa kadar ilerliyoruz özgeçmiş başlığına kadar olan yerleri listenin içine ekliyoruz
def getCitations(paragraphs):
    kontrol = 0
    kaynaklar = []
    for p in paragraphs:
        if p.text =='KAYNAKLAR':
            kontrol  =1
           
            continue
        elif p.text == 'ÖZGEÇMİŞ':
            break
        if kontrol ==1:
            text = p.text
            kaynaklar.append(text)
    return kaynaklar

#Özet başlığının olduğu yere kadar birşey almıyoruz özet başlığı ile anahtar kelimeler arasını bir liste ekliyoruz, anahtar kelimeleri de başka bir stringle geri döndürüyoruz 
def getSummaryandKeywords(paragraphs):
    kontrol = 0
    ozet = []
    keywords = ''
    for p in paragraphs:
        if p.text =='ÖZET':
            kontrol  =1
            continue
        elif 'Anahtar kelimeler' in p.text:
            keywords = p.text
            break
        if kontrol ==1:
            text = p.text
            ozet.append(text)
    return ozet,keywords
#Belge meta verileri içindeki author attributunu geri döndürüyoruz
def getAuthor(document):
    core_properties = document.core_properties
    author = core_properties.author
    title = core_properties.title
    return author,title
#Dokümanı aldığımız fonksiyon
def getThesis(thesisUrl):
   
    document = Document(thesisUrl)

    return document
#Normal style ile yazılmış paragrafların fontunu alıyoruz
def get_font(document):
    return document.styles['Normal'].font

#Başlık style ı ile yazılmış paragrafları alıyoruz. İki farklı başlığa odaklandık
def getImportantParagraphs(paragraphs,necessarry_styles):
    i_paragraphs = []
    for i in necessarry_styles:
        sub_i_p = []
        for p in paragraphs:
            if p.style == i:
                sub_i_p.append(p.text)

        i_paragraphs.append(sub_i_p)
    return i_paragraphs[0],i_paragraphs[1]

#Main fonksiyonu
def main():
    document = getThesis('Tez_SON.docx')
    author,title = getAuthor(document)
    font = get_font(document)
    paragraph = document.paragraphs
    ozet, keywords = getSummaryandKeywords(paragraph)
    citations = getCitations(paragraph)
    styles = document.styles
    stls = ['Heading 1', 'Heading 2', 'Heading 3', 'Heading 4','Heading 5','Heading 6','Normal Table','Title','TOC Heading','Header','table of figures']
    necessarry_styles = getNecessarryStyles(styles,stls)
    nec_title, main_title = getImportantParagraphs(paragraph,necessarry_styles)
    keywords = keywords[6:len(keywords)]
    keywords = keywords.split(',')
    data = {}
    data['author'] = author
    data['title'] = title
    data['font'] = font.name
    data['summary']=ozet
    data['citation']=citations
    data['forced_title'] = nec_title
    data['section_titles'] = main_title
    data['keywords']=keywords
    with open('data.txt', 'w',encoding='utf8') as outfile:
        json.dump(data, outfile,ensure_ascii=False)
main()