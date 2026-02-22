# Excel 转图片（Android）

把微信里下载的 Excel 表格（优先 `.xlsx`）转换成便于发布/分享的图片。

> 说明：由于使用 Apache POI 解析 Excel，本项目 `minSdk=26`（Android 8.0+）。

## 使用方式（目标体验）

1. 在微信里点开已下载的 Excel 文件 -> **用其他应用打开** -> 选择本应用
2. 预览 -> 导出图片 -> 直接分享微信 / 保存到相册
3. （可选）标注涂鸦：在应用内放大查看后做标记，支持导出“标注图片 / 标注PDF（矢量渲染，体积更小）”

## GitHub Releases 分发（给少数人安装）

- 用户只需要在 GitHub 仓库点 `Watch -> Releases`，就能只接收发版通知
- 每次发版会生成一个带版本号的 APK（例如：`excel2image-v0.1.10.apk`，点开即可安装）

## 发版（自动构建 APK）

本仓库已配置 GitHub Actions：当你 **发布 Release** 时自动构建 APK 并上传到 Release 的 Assets（文件名含版本号）。

推荐流程：

1) 更新版本号：修改 `app/build.gradle.kts` 里的 `versionName` / `versionCode`
2) 打 tag 并推送：

```bash
git tag -a v0.1.10 -m "v0.1.10"
git push origin v0.1.10
```

3) GitHub -> Releases -> Draft a new release
   - 选择 tag：`v0.1.10`
   - Publish release
4) 等 Actions 跑完，在 Release -> Assets 下载 `excel2image-v0.1.10.apk` 安装到手机

## 打包 APK

Debug（开发/自测）：

```bash
./gradlew :app:assembleDebug
```

Release（发给别人安装，建议用你自己的签名证书）：

```bash
./gradlew :app:assembleRelease
```

产物位置：

- `app/build/outputs/apk/debug/app-debug.apk`
- `app/build/outputs/apk/release/app-release.apk`

## Release 签名（强烈建议）

为什么需要签名？

- Android 只允许“同一签名”的 APK 覆盖更新；签名不一致会导致安装失败，必须先卸载旧版再装新版
- 你把 APK 发给同事/朋友自用时，稳定签名能保证后续更新不折腾

### 1) 生成 keystore（只做一次）

在项目根目录执行（示例）：

```bash
keytool -genkeypair -v -keystore excel2image.jks -alias excel2image -keyalg RSA -keysize 2048 -validity 36500
```

重要：`excel2image.jks` 和口令一定要备份好（丢了就基本无法给已安装用户“覆盖更新”）。

### 2) 本地打包签名（可选）

在项目根目录创建 `keystore.properties`（不要提交到 git）：

```properties
storeFile=excel2image.jks
storePassword=你的密码
keyAlias=excel2image
keyPassword=你的密码
```

### 3) GitHub Actions 签名（推荐，用于自动 Release）

在 GitHub 仓库里配置 Actions Secrets（Settings -> Secrets and variables -> Actions）：

- `ANDROID_KEYSTORE_BASE64`：把 `excel2image.jks` 做 base64 后粘贴进去（macOS 示例）
  ```bash
  base64 < excel2image.jks | tr -d '\n' | pbcopy
  ```
- `ANDROID_KEYSTORE_PASSWORD`
- `ANDROID_KEY_ALIAS`（一般是 `excel2image`）
- `ANDROID_KEY_PASSWORD`

如果不配置这些 secrets，Actions 会生成“临时 debug keystore”来签名：能安装，但每次签名可能不同，手机上更新时可能需要先卸载旧版。

提示：如果你曾经安装过“临时 debug keystore”签名的 Release 包，后续切换到你自己的 keystore 之后，**第一次升级通常需要卸载旧版再安装**（因为签名已经变了）。

### 4) 如何验证签名是否生效（每次更新都建议做）

- 手机上的验证方法：先安装上一个 Release 的 APK，再安装新 Release 的 APK
  - 能直接覆盖升级：签名 OK（同时 `versionCode` 也要递增）
  - 提示“签名不一致/无法安装”：说明签名没用上（或装的是不同包名/不同签名）

## 开发环境

- Android Studio（推荐）
- JDK 17+（本机已安装也可以，但 Android Gradle Plugin 可能要求 17）

## 本地复现/回归（导出 PDF 链路）

当你遇到“某些格子已换行，但行高没撑开导致文字被裁切”等问题时，建议把那份 Excel 放在 `.ys_files/`（不会提交到 git），然后用脚本在模拟器/真机上自动跑一次“导出 PDF”的链路：

```bash
scripts/local_repro_export_pdf.sh ".ys_files/outputs/0221-初步/excel文件/未结算送审1006个单项明细.xlsx"
```

脚本会把文件推到设备的应用外部目录，跑一次 `connectedDebugAndroidTest`，并把导出的 PDF 拉回到：

- `.ys_files/temp/repro-out.pdf`
