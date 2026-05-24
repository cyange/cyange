# 库存流水对比工具

本地 Web 工具，用于对比 SAP `EXPORT.XLSX` 与逻辑仓库存变动记录，按五组业务规则输出差异并保存 CSV。

## 启动

```bash
pip install -r requirements.txt
PORT=8766 python3 app.py
```

打开：

```text
http://127.0.0.1:8766/
```

## 说明

- 上传文件会临时保存到 `uploads/`，不会提交到 Git。
- 导出的 CSV 会保存到 `exports/`，不会提交到 Git。
- 工具只在本机运行，Excel 数据不会上传到外部服务。

## Cloudflare Pages

部署到 Cloudflare Pages 时，构建输出目录可以填写 `static`，也可以留空使用仓库根目录。线上版本在浏览器中解析 Excel，不依赖 Python 后端。

Excel 解析库 `static/xlsx.full.min.js` 已随项目一起部署，不依赖外部 CDN。
