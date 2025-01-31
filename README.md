# Word Cloud 词云生成系统

运行 `Setup.bat` 将自动检测并安装所需的依赖项，并生成 `Cloud.exe` 可执行文件。如果您不希望安装依赖项，或者当前正在进行重要任务，请谨慎执行 `Setup.bat`。如果您不希望安装依赖项但仍希望使用词云生成系统，可以直接使用我们提供的 `Cloud.exe` 可执行文件。

自动生成的 `Cloud.exe` 文件将位于当前目录下。首次使用时，加载过程可能会稍慢。

如果在运行 `Setup.bat` 时遇到错误，请仔细阅读错误提示，并根据提示进行相应操作。完成操作后，请重新运行 `Setup.bat`。

如果您的电脑已安装 Anaconda，可能需要执行 `conda remove pathlib`，但这可能会影响您的现有系统配置。如果您不希望更改系统配置，可以直接使用我们提供的 `Cloud.exe` 可执行文件。

如果您希望通过 Python 解释器直接运行 `Cloud.py`，可能会因路径问题导致运行失败。如果您坚持这样做，请阅读源码并修改相关路径。

我们不推荐直接使用 Python 解释器运行 `Cloud.py`，建议优先使用我们提供的 `Cloud.exe` 可执行文件，其次是自动生成的 `Cloud.exe`。

每次生成的词云图都会覆盖当前目录下的 `wordcloud.png` 文件。

**版权声明**

author：He Shuanglong
email：heshuanglong@e.gzhu.edu.cn
