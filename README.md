## 项目概述

PGA可以根据您上传的文本素材通过qwen-plus总结自动生成PPT，内嵌四套模板，支持上传其他模板。

本项目可以拆解为AI总结和PPT生成这两个部分，后续会在一定时间内更新和维护。

## 系统截图

## 使用说明

1. 将.env.example复制为.env

FLASK_DEBUG=1 # 调试模式  
DASHSCOPE_API_KEY=xxxx # 改成自己的API_KEY，百炼有免费额度

1. 定义模板在 @templates/ 里上传，仿制其他任意一套模板即可（每种不同占位符的页面顺序是固定的）

