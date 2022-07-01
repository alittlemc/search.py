# search.py
Python script to quickly query the contents of all word and Excel files in the directory.

> Windows环境下使用

# 依赖
> openpyxl、difflib、ctypes


# 必选:
        
        [str:关键字]
        最后一个输入参数为关键字,用于匹配,和-s作用相同,只有在无-s参数时生效.

        -s [str:关键字]
            -s选择关键字,使用最后一行.

# 可选*:
        -m [int:1-2]
            -m选择模式,1为word,2为excel,默认为2,为了确保兼容性,只扫描docx和xlsx文件.

        *-h  None
            显示帮助,输入此项后只输出提示.

        *-d <str:目录或文件,?={os.getcwd()}>
            需要查找的目录/文件,默认当前目录及子目录所有的word和excel文件.(可拓展.xlsm,.xltx,.xltm,不过不推荐)

        *-i <float:数字,?=0.6>
            excel文件模糊匹配相似度,0.1-1,对应为10%到100%,越大要求相似度越高,默认为0.6.
            word文件只能强匹配,只要段落内包含关键字即可.

        *-o
            除了终端打印额外输出csv文件,文件名为alittlemc_out.csv.

        *-a <int:id>
            test测试.

  ##  by alittlemc
  ##  version:    1.2(2022-05-10=15:50:44)
