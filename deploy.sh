#!/bin/bash
echo "开始部署到Vercel..."

# 安装Vercel CLI
npm install -g vercel

# 部署到Vercel
vercel --prod

echo "部署成功！"