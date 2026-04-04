# PPT Assistant

一个基于 VSTO 的 PowerPoint 插件，用来快速做玻璃质感效果。✨

## 功能

- `Transparent Glass`：把普通图形做成透明玻璃
- `White Glass`：做浅色玻璃
- `Black Glass`：做深色玻璃
- `Background`：把选中的图片铺满幻灯片，并生成模糊背景
- `Debug Apply`：用弹窗帮助定位卡死或报错位置

## 运行

1. 用 Visual Studio 打开 [PowerPointAddIn.slnx](/E:/OneDrive/计算机/自编程序/PowerPointAddIn/PowerPointAddIn.slnx)
2. 确保安装了桌面版 PowerPoint 和 VSTO 开发环境
3. 直接按 `F5`

## 原理

- 玻璃效果：背景填充 + 阴影 + 三维棱台
- 白/黑玻璃：先复制一层半透明颜色层，再组合后统一加效果
- 背景模糊：导出整页图片后做高斯近似模糊，再回写到幻灯片背景

## 说明

- 目前只建议对普通图形使用玻璃按钮
- PowerPoint 对“真渐变边框线”的代码支持很弱，所以当前版本走稳定优先路线 🛠️
