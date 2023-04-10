import scrapy
import sys
from scrapy.crawler import CrawlerProcess
import datetime

class TGEgasSpider(scrapy.Spider):
    name = 'TGEgas'
    allowed_domains = ['www.tge.pl']

    def __init__(self, *args, **kwargs):
        super(TGEgasSpider, self).__init__(*args, **kwargs)
        self.date = datetime.date.today()
        self.date -= datetime.timedelta(days = 8)
        self.day = self.date.day
        self.month = self.date.month
        self.year = self.date.year
    
    def start_requests(self):
        link = ""
        for i in range(7):
            link = (f'https://tge.pl/gaz-rdn?dateShow={self.day}-{self.month}-{self.year}')  
            self.date += datetime.timedelta(days=1)
            (self.day,self.month,self.year) =  self.date.day,self.date.month,self.date.year

            yield scrapy.Request(url = link,callback=self.parse, headers={
                'User-Agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.131 Safari/537.36'
            })
    def parse(self, response):
        price = float(response.xpath("(//tr/td[@class='footable-visible'])[1]/text()").get().replace(",","."))   
    
        yield {
            "price": price,
        }

if __name__ == "__main__":
    try:
        LOG = bool(int(sys.argv[1]))
    except:
        LOG = 0
    spider_settings = {
        "FEEDS": {
        "TGEgas.json": {"format": "json"},}, 
         "CONCURRENT_REQUESTS" : 1,
         "LOG_ENABLED" : LOG}
        
    process = CrawlerProcess(spider_settings)
    process.crawl(TGEgasSpider)
    process.start()
 
