import scrapy
import sys
from scrapy.crawler import CrawlerProcess
import datetime

class BASEgasSpider(scrapy.Spider):
    name = 'BASEgas'
    allowed_domains = ['www.tge.pl']
    def __init__(self, *args, **kwargs):
        super(BASEgasSpider, self).__init__(*args, **kwargs)
        self.date = datetime.date.today()
        self.date -= datetime.timedelta(days = 7)
        self.day = self.date.day
        self.month = self.date.month
        self.year = self.date.year
    
    def start_requests(self):
        link = ""
        for _ in range(5):
            link = (f"https://tge.pl/gaz-otf?dateShow={str(self.day).rjust(2,'0')}-{str(self.month).rjust(2,'0')}-{self.year}")  
            self.date += datetime.timedelta(days=1)
            (self.day,self.month,self.year) = self.date.day, self.date.month, self.date.year 

            yield scrapy.Request(url = link,callback=self.parse, headers={
                'User-Agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.131 Safari/537.36'
            })
        #(//div[@class='table-responsive wyniki-footable-kontrakty-terminowe-0']//tr[@class='period-'])
    def parse(self, response):
        rows = response.xpath("(//div[@class='table-responsive wyniki-footable-kontrakty-terminowe-0'])")
        DKR = []
        MWh = []
        try:
            for i in [15,16]:
                DKR.append(float(rows.xpath(f"((//tr[@class='period-'])[{i}]/td)[4]/text()").get().replace(',','.').replace('-','0')))
                MWh.append(float(rows.xpath(f"((//tr[@class='period-'])[{i}]/td)[7]/text()").get().replace(',','.').replace('-','0')))
            yield {
                "BASE25-DKR": DKR[0],
                "BASE25-MWh" : MWh[0],
                "BASE26-DKR": DKR[1],
                "BASE26-MWh" : MWh[1],
            }
        except:
            print("\nSearch did not find anything")

if __name__ == "__main__":
    try:
        LOG = bool(int(sys.argv[1]))
    except:
        LOG = 0
    spider_settings = {
            "FEEDS": {
                    "BASEgas.json": 
                        {"format": "json"},}, 
            "CONCURRENT_REQUESTS" : 1,
            "LOG_ENABLED" : LOG}

    process = CrawlerProcess(spider_settings)
    process.crawl(BASEgasSpider)
    process.start()

