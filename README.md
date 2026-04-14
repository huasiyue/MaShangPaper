# MaShangPaper

## 后端启动

```bash
cd /Users/yun/Documents/project/trea_project/github/MaShangPaper/backend
/opt/anaconda3/envs/pytorch/bin/python -m pip install -r requirements.txt
/opt/anaconda3/envs/pytorch/bin/python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

## 前端启动

```bash
cd /Users/yun/Documents/project/trea_project/github/MaShangPaper/frontend
npm install
npm run dev
```

## 图片插入

```md
![图片标题|width=8cm|align=center](http://127.0.0.1:8000/api/assets/资源ID)
```

支持：

- `width=8cm`
- `align=center`
- `align=left`
- `align=right`
