import scrapy
import sys
from scrapy.crawler import CrawlerProcess
import datetime


class GPISpider(scrapy.Spider):
    name = 'GPI'
    allowed_domains = ["www.gpi.tge.pl"]

    
    def __init__(self, *args, **kwargs):
        super(GPISpider, self).__init__(*args, **kwargs)
        self.date = datetime.date.today()
        self.date -= datetime.timedelta(days = 7)
        self.day = self.date.day
        self.month = self.date.month
        self.year = self.date.year

    def start_requests(self):
        link = (f"http://gpi.tge.pl/pl/zestawienie-ubytkow?p_p_id=gpicalendar_WAR_gpicalendarportlet&p_p_lifecycle=0&p_p_state=normal&p_p_mode=view&p_p_col_id=column-1&p_p_col_count=1&_gpicalendar_WAR_gpicalendarportlet_direction=0&_gpicalendar_WAR_gpicalendarportlet_current={self.year}-{self.month}-{self.day}")
        yield scrapy.Request(url = link, callback=self.parse, headers={
            'User-Agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.131 Safari/537.36'
            })
    def parse(self, response):  
        planned = []
        notplanned = []
        for ubytek in response.xpath("(//tr[@class='summary'])[1]/td[@class!='colSummary']"):
            planned.append(int(ubytek.xpath(".//text()").get()))
        for ubytek in response.xpath("(//tr[@class='summary'])[2]/td[@class!='colSummary']"):
            notplanned.append(int(ubytek.xpath(".//text()").get()))      
        summaric_leaks = list(map(lambda x,y : x+y,planned,notplanned))
        yield {
            "planned" : planned,
            "notplanned" : notplanned,
            "summaric_leaks" : summaric_leaks
        }


if __name__ == "__main__":
    try:
        LOG = bool(int(sys.argv[1]))
    except:
        LOG = 0
    spider_settings = {
            "FEEDS": {
            "GPI.json": {"format": "json"},}, 
             "CONCURRENT_REQUESTS" : 1,
             "LOG_ENABLED" : LOG}

    process = CrawlerProcess(spider_settings)
    process.crawl(GPISpider)
    process.start()
 
