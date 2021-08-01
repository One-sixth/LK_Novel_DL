# LK_Novel_DL
 轻国小说下载到Docx (python)  

```txt

LK Novel Download Tool
LK小说下载工具

示例
python3 lk_novel_dl.py https://www.lightnovel.us/cn/detail/495565

下载完成后，输出文件在当前目录下，默认名字为 out.docx

-------------------------------

usage: lk_novel_dl.py [-h] [--no-title] [--no-cache]
                      [--replace-txt REPLACE_TXT] [--help-replace-txt]
                      [--out OUT]
                      url

positional arguments:
  url                   这是填入轻国小说的url

optional arguments:
  -h, --help            show this help message and exit
  --no-title            设定该标志可关闭自动的标题检测和标题生成
  --no-cache            设定该标志可以关闭缓存，将会始终使用在线链接
  --replace-txt REPLACE_TXT
                        高级用法，替换url指向资源，使用方法参见 --help-replace-txt
  --help-replace-txt    显示 --replace-txt 怎么用
  --out OUT             输出文件的路径

```

# 依赖
```txt
opencv-python
imageio
numpy
beautifulsoup4
requests
python-docx
```

# 怎么用

非常简单  
```sh
python3 lk_novel_dl.py https://www.lightnovel.us/cn/detail/495565
```

下载完成后，输出文件在当前目录下，默认名字为 out.docx  


# 高级用法

## replace_txt

本功能用来替换指定url为本地资源。  
某些轻国小说页面结构有问题，例如 https://www.lightnovel.us/cn/detail/336435  
需要把网页下载下来，然后手动修正html结构，不然会缺失某些项。  
replace-txt 是一个普通的文本文件，可以用notepad编辑和生成。  
文件结构为 行式结构，行间使用回车符分隔，行内使用空格符分隔。  
第一个列为请求的url，第二列为要替换的本地资源的路径。  

replace-txt 结构  
```
请求的url1 本地资源路径1
请求的url2 本地资源路径2
```

例子，参见 example-replace-txt.txt  
```
https://www.lightnovel.us/cn/detail/336435 336435_fix.html
https://www.lightnovel.us/cn/detail/392689 392689_fix.html
```