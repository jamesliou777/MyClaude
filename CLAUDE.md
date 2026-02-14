# 專案記憶

## 專案概述
- 單頁式工程計算器（`index.html`），部署於 GitHub Pages
- GitHub 帳號：`jamesliou777`
- 儲存庫：`MyClaude`，分支：`main`
- 最新提交：`723f2c1` - 新增 RC 梁撓曲設計計算器（ACI 318）

## 架構
- **單一 HTML 檔案**，內嵌 CSS 與 JS
- 分頁導覽：beam、column、concrete、rebar、rcbeam、unit
- 頁首設有全域單位系統選擇器

## 單位系統
- 支援：kN-m、Kgf-cm、tf-m、N-mm、tf-cm、Kgf-m
- 架構：`factorsToSI`、`unitSystemDefs`、`fieldUnits`（含 `dim` 屬性）
- 函式：`switchUnitSystem()`、`resultInSystem()`
- RC 梁以 kgf/cm² 為應力基底，tf·m 為彎矩基底

## 功能
- RC 梁撓曲設計（ACI 318）：Rn 法、β1、ρmin/ρmax、εt 應變檢核
- 已移除區段：電機計算、邊坡穩定、管路水路

## 部署
- 從 main 分支部署至 GitHub Pages
- 需手動至 Settings → Pages 啟用（若尚未啟用）

## 待處理 / 暫緩
- 程式碼保護（混淆、WASM、後端 API）— 使用者暫緩，仍在考慮中
- GitHub Pages 手動啟用狀態未確認

## 提交歷史
- `bd9fb75` 初始提交：新增 index.html
- `2d16cae` 新增全域單位系統切換器
- `86d027b` 移除電機計算、邊坡穩定、管路水路區段
- `723f2c1` 新增 RC 梁撓曲設計計算器（ACI 318）
