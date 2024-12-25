# iOS 实况照片 -> 小米动态照片 转换脚本

本工具可以将heic/jpeg + mov格式的iOS实况照片转换为小米相册可以识别的动态照片格式。


# 脚本使用方法（命令行版）：
```
git clone https://github.com/Serendo/LivePhoto2XiaomiPhoto.git
cd LivePhoto2XiaomiPhoto
. .\converter-cli.ps1
# 以转换'C:\Users\Serendo\Downloads\TE ST\' 这个目录中的实况照片为例
Convert-LivePhotoFolder 'C:\Users\Serendo\Downloads\TE ST\'
```

# 图形版
```
git clone https://github.com/Serendo/LivePhoto2XiaomiPhoto.git
cd LivePhoto2XiaomiPhoto
.\converter-gui.ps1
```

完整的换机照片迁移方案可以参考这里[B站](https://www.bilibili.com/opus/1006443152534405127)

# 感谢
1. exiftool
2. ffmpeg
