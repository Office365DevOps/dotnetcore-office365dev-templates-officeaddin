# Office Add-in Toolbox 

> Ares Chen @ 2019-3-3

This is a toolbox for Office Add-in developers, It will be develop and update continuely.

## How to Install

You can install this toolbox via `dotnet tool install --global dotnetcore-officeaddin-toolbox`

## How to Use

Currently, this toolbox can only support Windows platform, and only one most popular scenario (`sideload`) was supported.

## Use sideload command

Please run `office-toolbox sideload manifest.xml Excel` to sideload your Add-in in second, this command have three parameters as below

1. office-toolbox, this is the tool name, you always use this name for several commands.
1. sideload, this is the command name,currently we only support `sideload` command.
1. manifest.xml, this is the manifest file of your addin project.
1. Excel, this is your hosted application name, currently we support `Excel`,`Word`,`PowerPoint`.

Happy coding!

## Office Add-in toolbox 工具集中文介绍

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