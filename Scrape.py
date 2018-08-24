# -*- coding: utf-8 -*-
"""
Created on Sun Jun 24 14:59:11 2018

@author: Anurag
"""

# import libraries
import urllib.request
from bs4 import BeautifulSoup
import re
from datetime import datetime
import xlsxwriter

class AppURLopener(urllib.request.FancyURLopener):
    version = "Mozilla/5.0"

class WebScrapper:
    def __init__(self, homePage, rootUrl):
        self.homePage = homePage
        self.rootUrl = rootUrl
        self.currentUrl = self.rootUrl
        self.opener = AppURLopener()
        self.links = []


    def findTags(self, url, tag, tagDict):
        # query the website and return the html to the variable ‘page’
        # parse the html using beautiful soup and store in variable `soup`
        page = self.opener.open(url)
        #print("warn", url)
        soup = BeautifulSoup(page, 'html.parser')
        #print("warn")

        # find all the tags
        scrappedTags = soup.find_all(tag, attrs=tagDict)
        #print("warn")
        #print(scrappedTags)
        return scrappedTags



    def getAllAnchorTags(self, scrappedTags, regExp):
        for eachBlock in scrappedTags:
            for eachTag in eachBlock:
                matchObj = re.match(r'(.*)(\/shareprice.*)(")', str(eachTag), re.M | re.I)
                try:
                    relativeDir = matchObj.group(2)
                    completeLink = self.homePage + relativeDir
                    #print(completeLink)
                    self.links.append([eachBlock.text.strip().split("\n"), completeLink])
                except:
                    #print("error in regex")
                    pass
        return self.links

    def scrapeLink(self, tags):
        for each in tags:
            name = each.text.strip().split()  # strip() is used to remove starting and trailing
            name = ''.join(name)
            # print(name)
            matchObj = re.match(r'.*(Bidprice)(\d*\.\d*)(Openprice)(\d*\.\d*)(Askprice)(\d*\.\d*)(Prevclose)(\d*\.\d*).*',
                                str(name), re.M | re.I)
            try:
                #print(matchObj.group(1), matchObj.group(2), matchObj.group(3), matchObj.group(4))
                return matchObj.group(2), matchObj.group(4), matchObj.group(6), matchObj.group(8)
            except:
                pass


class excelWriter:
    def __init__(self, fileName):
        self.saveFile = fileName
        self.workbook = xlsxwriter.Workbook(self.saveFile)
        self.worksheet = self.workbook.add_worksheet()
        self.worksheet.write('A1', "Overview")
        self.worksheet.write('B1', "Full Name")
        self.worksheet.write('C1', "Years In Practice")
        self.worksheet.write('D1', "Language")
        self.worksheet.write('E1', "Office Location")
        self.worksheet.write('F1', "Hospital Affiliation")
        self.worksheet.write('G1', "Specialities and Sub Specialities")
	self.worksheet.write('H1', "Education and Medical Training")
	self.worksheet.write('I1', "Certification and Licensure")
        self.currentRowCounter = 2

    def writeIntoFile(self, content):
        for eachEntry in content:
            self.worksheet.write('A'+str(self.currentRowCounter), str(eachEntry[0]))
            self.worksheet.write('B'+str(self.currentRowCounter), str(eachEntry[1]))
            self.worksheet.write('C'+str(self.currentRowCounter), str(eachEntry[2]))
            self.worksheet.write('D'+str(self.currentRowCounter), str(eachEntry[3]))
            self.worksheet.write('E'+str(self.currentRowCounter), str(eachEntry[4]))
            self.worksheet.write('F'+str(self.currentRowCounter), str(eachEntry[5]))
            self.worksheet.write('G'+str(self.currentRowCounter), str(eachEntry[6]))
	    self.worksheet.write('H'+str(self.currentRowCounter), str(eachEntry[7]))
	    self.worksheet.write('I'+str(self.currentRowCounter), str(eachEntry[8]))
            self.currentRowCounter += 1

    def closeExcel(self):
        self.workbook.close()




def main():
    dataList = []
    webObj = WebScrapper("https://health.usnews.com", "https://health.usnews.com/doctors/specialists-index/new-jersey")
    tags = webObj.findTags("https://health.usnews.com/doctors/specialists-index/new-jersey",'tr' ,{'class': 'stdTblRow'})
    links = webObj.getAllAnchorTags(tags, '(.*)(\/shareprice.*)(")')
    #print(links)
    count = 1
    for eachLink in links:
        try:
            tags = webObj.findTags(eachLink[1], 'div', {'class': 'ui-helper-clearfix'})
            bid, open, ask, prev = webObj.scrapeLink(tags)
            currentData = [eachLink[0][0], eachLink[0][1], eachLink[0][2], bid, open, ask, prev]
            dataList.append(currentData)
            print(currentData)
            count+=1

        except:
            print("error at", count)

    excel = excelWriter("finalExcelFile.xlsx")
    excel.writeIntoFile(dataList)
    excel.closeExcel()
    
    

    print(count)

main()
print("Finished")