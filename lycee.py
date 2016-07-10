from bs4 import BeautifulSoup as bs
import urllib2
import re
import xlwt
# import uniout

category = ["CH","EV","IT","AR"]
page = 0

final_CHpage = 93
final_EVpage = 8
final_ITpage = 3
final_ARpage = 2

AllCards = []
CHCardsList = []
CHCards = []
while (page < final_CHpage + 1 ):
        base_url="http://adlyrs.ddo.jp/lycee/?category="
        url = base_url + str(category[0]) + "&p=" + str(page)
        html = urllib2.urlopen(url).read()
        soup = bs(urllib2.urlopen(url), "html.parser")

        card_list = soup.find_all("div", {"class":"card"})
        for card in card_list:
                number = card.find("a").string
                title = card.find("span", {"class":"Hint"}).string
                ex = title.findNext("td").string
                cost = ex.findNext("td").string
                field = card.find("td", {"class":"field"})
                AP = field.findNext("td").string
                DP = AP.findNext("td").string
                SP = DP.findNext("td").string
                Sex = SP.findNext("td").string
                btext_obj = card.find("p", {"class": "btext"})
                if btext_obj :
                        btext = btext_obj.string
                else:
                        btext = ""
                text_obj = card.find("p", {"class": "text"})
                texts = ''
                if text_obj :
                        texts = text_obj.contents[0]
                                # texts = texts + text.string
                else:
                        texts = ""
                version = text_obj.findNext("td").string
                source = version.findNext("td").string
                rare = source.findNext("td").string
                illustration = source.findNext("a").string
                CHCards.append(number)
                CHCards.append(title)
                CHCards.append(ex)
                CHCards.append(cost)
                CHCards.append(str(field))
                CHCards.append(AP)
                CHCards.append(DP)
                CHCards.append(SP)
                CHCards.append(Sex)
                CHCards.append(btext)
                CHCards.append(texts)
                CHCards.append(version)
                CHCards.append(source)
                CHCards.append(rare)
                CHCards.append(illustration)
                CHCardsList.append(CHCards)
                CHCards = []
        page = page + 1
        print page

book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Character")
n = len(CHCardsList)

i = 0
j = 0
while i < n:
        while j < 15:
                sheet1.write(i+1, j, CHCardsList[i][j])
                j = j + 1
        j = 0
        i = i+1
book.save("lycee.xls")
print "Ch done"

page = 0
EVCardsList = []
EVCards = []


while (page < final_EVpage + 1 ):
        base_url="http://adlyrs.ddo.jp/lycee/?category="
        url = base_url + str(category[1]) + "&p=" + str(page)
        html = urllib2.urlopen(url).read()
        soup = bs(urllib2.urlopen(url), "html.parser")

        card_list = soup.find_all("div", {"class":"card"})
        for card in card_list:
                number = card.find("a").string
                title = card.find("span", {"class":"Hint"}).string
                ex = title.findNext("td").string
                cost = ex.findNext("td").string
                field = card.find("td", {"class":"field"})
                text_obj = card.find("p", {"class": "text"})
                texts = ''
                if text_obj :
                        texts = text_obj.contents[0]
                                # texts = texts + text.string
                else:
                        texts = ""
                version = text_obj.findNext("td").string
                source = version.findNext("td").string
                rare = source.findNext("td").string
                illustration = source.findNext("a").string
                EVCards.append(number)
                EVCards.append(title)
                EVCards.append(ex)
                EVCards.append(cost)
                EVCards.append(texts)
                EVCards.append(version)
                EVCards.append(source)
                EVCards.append(rare)
                EVCards.append(illustration)
                EVCardsList.append(EVCards)
                EVCards = []
        page = page + 1
        print page

sheet2 = book.add_sheet("Event")
n = len(EVCardsList)

i = 0
j = 0
while i < n:
        while j < 9:
                sheet2.write(i+1, j, EVCardsList[i][j])
                j = j + 1
        j = 0
        i = i+1
book.save("lycee.xls")
print "Ev done"

page = 0
ITCardsList = []
ITCards = []


while (page < final_ITpage + 1 ):
        base_url="http://adlyrs.ddo.jp/lycee/?category="
        url = base_url + str(category[2]) + "&p=" + str(page)
        html = urllib2.urlopen(url).read()
        soup = bs(urllib2.urlopen(url), "html.parser")

        card_list = soup.find_all("div", {"class":"card"})
        for card in card_list:
                number = card.find("a").string
                title = card.find("span", {"class":"Hint"}).string
                ex = title.findNext("td").string
                cost = ex.findNext("td").string
                field = card.find("td", {"class":"field"})
                text_obj = card.find("p", {"class": "text"})
                texts = ''
                if text_obj :
                        texts = text_obj.contents[0]
                                # texts = texts + text.string
                else:
                        texts = ""
                version = text_obj.findNext("td").string
                source = version.findNext("td").string
                rare = source.findNext("td").string
                illustration = source.findNext("a").string
                ITCards.append(number)
                ITCards.append(title)
                ITCards.append(ex)
                ITCards.append(cost)
                ITCards.append(texts)
                ITCards.append(version)
                ITCards.append(source)
                ITCards.append(rare)
                ITCards.append(illustration)
                ITCardsList.append(ITCards)
                ITCards = []
        page = page + 1
        print page

sheet3 = book.add_sheet("Item")
n = len(ITCardsList)

i = 0
j = 0
while i < n:
        while j < 9:
                sheet3.write(i+1, j, ITCardsList[i][j])
                j = j + 1
        j = 0
        i = i+1
book.save("lycee.xls")
print "It done"

page = 0
ARCardsList = []
ARCards = []


while (page < final_ARpage + 1 ):
        base_url="http://adlyrs.ddo.jp/lycee/?category="
        url = base_url + str(category[3]) + "&p=" + str(page)
        html = urllib2.urlopen(url).read()
        soup = bs(urllib2.urlopen(url), "html.parser")

        card_list = soup.find_all("div", {"class":"card"})
        for card in card_list:
                number = card.find("a").string
                title = card.find("span", {"class":"Hint"}).string
                ex = title.findNext("td").string
                cost = ex.findNext("td").string
                field = card.find("td", {"class":"field"})
                text_obj = card.find("p", {"class": "text"})
                texts = ''
                if text_obj :
                        texts = text_obj.contents[0]
                                # texts = texts + text.string
                else:
                        texts = ""
                version = text_obj.findNext("td").string
                source = version.findNext("td").string
                rare = source.findNext("td").string
                illustration = source.findNext("a").string
                ARCards.append(number)
                ARCards.append(title)
                ARCards.append(ex)
                ARCards.append(cost)
                ARCards.append(texts)
                ARCards.append(version)
                ARCards.append(source)
                ARCards.append(rare)
                ARCards.append(illustration)
                ARCardsList.append(ARCards)
                ARCards = []
        page = page + 1
        print page

sheet4 = book.add_sheet("Area")
n = len(ARCardsList)

i = 0
j = 0
while i < n:
        while j < 9:
                sheet4.write(i+1, j, ARCardsList[i][j])
                j = j + 1
        j = 0
        i = i+1
book.save("lycee.xls")
print "Ar done"