import time
from lxml import etree
import aiohttp
import asyncio
import xlrd
import xlwt

skuId_off_sale = []
sem = asyncio.Semaphore(10) #信号量，控制协程数，防止爬的过快
header = {'Cookie':'OCSSID=4df0bjva6j7ejussu8al3eqo03','User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36''(KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36'}

async def get_status(skuId):
    async with sem:
        #async with是异步上下文管理器
        async with aiohttp.ClientSession() as session: #获取session
            async with session.request('GET',url='https://item.jd.com/'+str(skuId)+'.html',headers=header) as resp: #提出请求
                 html = await resp.read() #直接获取bytes
                 status = etree.HTML(html).xpath('/html/body/div[6]/div/div[2]/*[@class="itemover-tip"]/text()') #xpath解析获取html内容
                 if len(status):
                    # print('编号 %d 已下架！' % int(skuId))
                   skuId_off_sale.append(skuId)


def main(): 
    excel_data = xlrd.open_workbook('JD_goods.xls') #打开Excel文件
    table = excel_data.sheets()[0] #获取第一个工作表
    nrows = table.nrows #获取表中有效行
    JDskuIds = []
    i = 1
    while i < nrows:
        JDskuIds.append(int(table.cell_value(i,0))) #获取单元格数据
        i = i + 1
    # print('此商品编码是：%s' % JDId)

    #调用爬虫
    start = time.time()
    loop = asyncio.get_event_loop() #获取事件循环
    tasks = [get_status(skuId) for skuId in JDskuIds] #任务列表
    loop.run_until_complete(asyncio.wait(tasks)) #激活协程
    loop.close() #关闭事件循环
    print('爬取总耗时：%.5f秒' %float(time.time()-start))

    #将下架的编号保存至新建xls表格
    book = xlwt.Workbook() #新建工作簿
    sheet = book.add_sheet('下架商品') #添加工作页
    sheet.write(0,0,'下架商品')
    n = 1
    for id in skuId_off_sale:
        sheet.write(n,0,id)
        n = n + 1
    book.save(filename_or_stream='下架商品.xls')


if __name__ == '__main__':
    main()