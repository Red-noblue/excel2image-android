# Excel 转图片（Android）

把微信里下载的 Excel 表格（优先 `.xlsx`）转换成便于发布/分享的图片。

## 使用方式（目标体验）

1. 在微信里点开已下载的 Excel 文件 -> **用其他应用打开** -> 选择本应用
2. 预览 -> 导出图片 -> 直接分享微信 / 保存到相册
3. （可选）标注涂鸦：在应用内放大查看后做标记，支持导出“标注图片 / 标注PDF”（PDF 分享更清晰）

## GitHub Releases 分发（给少数人安装）

- 用户只需要在 GitHub 仓库点 `Watch -> Releases`，就能只接收发版通知
- 每次发版上传一个 `universal` APK（点开即可安装）

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

## Release 签名（建议）

1) 生成 keystore（示例）：

```bash
keytool -genkeypair -v -keystore excel2image.jks -alias excel2image -keyalg RSA -keysize 2048 -validity 36500
```

2) 在项目根目录创建 `keystore.properties`（不要提交到 git）：

```properties
storeFile=excel2image.jks
storePassword=你的密码
keyAlias=excel2image
keyPassword=你的密码
```

## 开发环境

- Android Studio（推荐）
- JDK 17+（本机已安装也可以，但 Android Gradle Plugin 可能要求 17）
