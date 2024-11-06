# windows 环境配置

1. 安装 node.js
   浏览器访问 https://nodejs.org/zh-cn/download/package-manager，根据指引，安装 node.js

2. 安装 pnpm
   打开命令行，输入 `npm install -g pnpm`

3. 安装 pnpm 的依赖
   打开命令行，命令行中进入项目目录，输入 `pnpm i`

4. 运行项目
   进入项目目录，将需要处理的 excel 文件放在 origin 文件夹下
   打开命令行，进入项目目录，输入 `pnpm start`
   处理后的文件会放在 result 文件夹下

---

# 打包 windows 应用

确保安装了 pkg
进入项目目录，输入 `pkg .`
