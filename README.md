# Auto-display-impact-factor
The program can display a sci magazine' impact factor when you copy its name in windows clipboard.

这个项目源于在写论文时需要引用较高影响因子的论文，所以最好可以将某一个期刊的影响因子即时地显示出来。

使用时，将期刊名复制到剪贴板中，会自动显示。
一共三个版本：offline是离线版，主要是查数据表a.xls，online1基于phantomJS虚拟浏览器，online2采用向letpub发送请求的方式获取数据。推荐使用online2 。
