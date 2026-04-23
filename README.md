## 项目概述

ppt-generate-agent

**2026.3.25 V1.0**

根据您上传的文本素材通过qwen系列模型总结自动生成PPT，内嵌四套模板，支持上传其他模板。

本项目可以拆解为AI总结和PPT生成这两个部分，后续会在一定时间内更新和维护。

**2026.4.23 V2.0**

优化了生成PPT的美观性。

增加了字体会适配文本框进行缩小的功能。

修复了生成ppt背景丢失的bug。

## 系统截图

<img width="1372" height="1364" alt="a22a829fd2f7a3a788e9a7e490d0988b" src="https://github.com/user-attachments/assets/5589464d-33ba-4959-a857-9c04d62299df" />

<img width="1352" height="1440" alt="fdd77c39d18935cf3b0dbacdf3a5d08a" src="https://github.com/user-attachments/assets/31a03dc6-cb1d-4391-b4a9-b0c8f2874418" />

## 使用说明

1. 将.env.example复制为.env

FLASK_DEBUG=1 # 调试模式  

DASHSCOPE_API_KEY=xxxx # 改成自己的API_KEY，百炼有免费额度，https://bailian.console.aliyun.com

2. 定义模板在 @templates/ 里上传，仿制其他任意一套模板即可（每种不同占位符的页面顺序是固定的）

