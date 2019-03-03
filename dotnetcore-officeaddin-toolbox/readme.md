# Office Add-in Toolbox 工具集

> 陈希章 于 2019-3-3

这是一套提供给Office Add-in开发人员的工具集，将持续开发和更新。

## 如何安装

请通过 `dotnet tool install --global dotnetcore-officeaddin-toolbox` 进行安装和升级

## 如何使用

目前该工具仅支持Windows平台，而且仅支持一个命令（sideload）用来简化在本地进行开发调试的过程。

## sideload 命令

请在生成好的Add-in项目的目录，运行类似的命令行，`office-toolbox sideload manifest.xml Excel` 即可自动加载Add-in并且打开Office程序。这里的命令参数说明如下

1. office-toolbox 是工具名称
1. sideload 是命令（目前只有这一个命令）
1. manifest.xml 是清单文件，请提供相对路径即可
1. Excel是指应用程序，目前支持Excel，Word，PowerPoint