# Word文档格式优化项目

## 项目概述
帮助用户根据格式要求文件优化Word文档格式的自动化工具。

## 技术栈
- 后端: Python + FastAPI
- 前端: React + TypeScript
- Word处理: python-docx

## 快速开始

### 后端启动
```bash
cd backend
pip install -r requirements.txt
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

### 前端启动
```bash
cd frontend
npm install
npm start
```

## 项目结构
```
word-formatter/
├── backend/                 # FastAPI后端
│   ├── app/
│   │   ├── main.py         # FastAPI应用入口
│   │   ├── api/            # API路由
│   │   ├── services/       # 业务逻辑
│   │   └── models/         # 数据模型
│   └── requirements.txt
├── frontend/               # React前端
│   ├── src/
│   │   ├── components/     # 组件
│   │   ├── pages/          # 页面
│   │   └── services/       # API调用
│   └── package.json
└── README.md
```

## 主要功能
1. 文件上传 (原文档 + 格式要求文档)
2. Word文档解析和格式分析
3. 格式对比和优化建议
4. 自动格式应用和文档生成
5. 处理结果下载

## 开发状态
- ✅ 项目结构搭建完成
- ✅ 基础前后端框架
- ⏳ 文件上传功能 (待实现)
- ⏳ Word文档处理 (待实现)
- ⏳ 格式优化逻辑 (待实现)