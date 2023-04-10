import scrapy
import sys
from scrapy.crawler import CrawlerProcess
import datetime

class TGESpider(scrapy.Spider):
    name = 'TGE'
    allowed_domains = ['www.tge.pl']
    def __init__(self, *args, **kwargs):
        super(TGESpider, self).__init__(*args, **kwargs)
        self.date = datetime.date.today()
        self.date -= datetime.timedelta(days = 8)
        self.day = self.date.day
        self.month = self.date.month
        self.year = self.date.year
    
    def start_requests(self):
        link = ""
        for i in range(7):
            link = (f'https://tge.pl/energia-elektryczna-rdn-tge-base?date_start={self.year}-{self.month}-{self.day}')  
            self.date += datetime.timedelta(days=1)
            (self.day,self.miesiac,self.rok) = self.date.day, self.date.month, self.date.year 

            yield scrapy.Request(url = link,callback=self.parse, headers={
                'User-Agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.131 Safari/537.36'
            })
    def parse(self, response):
        price = float(response.xpath("(//tr/td[@class='footable-visible'])[1]/text()").get().replace(",","."))   
        min_price = 1000000.0
        max_price = 0.0
        max_hour = ""
        min_hour = ""
        for row in response.xpath("//table[@class='footable table table-hover table-padding']/tbody/tr"):
            hour = int(row.xpath("normalize-space(.//td/b/text())").get().split("-")[1])
            value = float(row.xpath("normalize-space(.//td[2]/text())").get().replace(",","."))
            if min_price > value:
                min_price = value
                min_hour = hour
            if max_price < value:
                max_price = value
                max_hour = hour
    
        yield {
            "price": price,
            "min price" : min_price,
            "hour for min": min_hour,
            "max price" : max_price,
            "hour for max" : max_hour
        }

if __name__ == "__main__":
    try:
         LOG = bool(int(sys.argv[1]))
    except:
         LOG = 0
    spider_settings = {
            "FEEDS": {
                    "TGEenergy.json": 
                        {"format": "json"},}, 
            "CONCURRENT_REQUESTS" : 1,
            "LOG_ENABLED" : LOG}

    process = CrawlerProcess(spider_settings)
    process.crawl(TGESpider)
    process.start()
 


