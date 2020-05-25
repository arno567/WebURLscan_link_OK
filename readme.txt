
1.识别xlsx每张sheet中的“网站域名|域名|URL|url|域名/URL|域名/url|链接”这些列的所有url（目前功能每张sheet只识别一列），检测的对象xlsx有格式要求，参考"xlsx格式要求及结果展示.jpg"。检测结果输出到url对应行的最后端。
2.若搜集过来的url量很大，不能确定采用的是http还是https，没关系 工具会对所有url进行http和https测试。（识别的url格式：http://xxxx/xx、https://xxx/xx、xxx.xxx/xxx/）
3.优点：检测准确率高，运用线程池检查速度较快；直接修改xlsx文件方便写报告。
4.缺点：有些网站通过https访问会跳转到防火墙或者vpn 识别结果也会是正常。
5.简单使用：根据脚本头信息安装对应的模块后-python3 web可用性检测工具.py
