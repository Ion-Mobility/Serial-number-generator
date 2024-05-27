步骤1：下载和安装PyCharm

首先，您需要从JetBrains官网下载并安装PyCharm。
步骤2：设置Python解释器
1.打开PyCharm，并创建或打开您的项目。
2.单击顶部菜单中的“File”（文件）。
3.选择“Settings”（设置）。
4.在设置窗口中，展开“Project: [Your Project Name]”并选择“Python Interpreter”（Python 解释器）。
5.单击右上角的齿轮图标，并选择“Add...”（添加...）。
6.在弹出的窗口中，选择您系统中已安装的Python解释器，然后单击“OK”。
步骤3：安装项目依赖
1.在您的项目根目录中创建一个名为requirements.txt的文件，如果已经存在则跳过此步骤。
2.将您项目所需的所有依赖项以及其版本号添加到requirements.txt文件中，每行一个依赖项，
3.在PyCharm的终端中，导航到您的项目根目录。
4.运行以下命令来安装所有依赖项：
pip install -r requirements.txt
