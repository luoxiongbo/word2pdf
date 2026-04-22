'use strict';

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const os = require('os');

/**
 * 创建一个带有正确字体度量映射的 LibreOffice 用户配置目录。
 */
function buildLoProfile(profileDir, sobinary) {
  if (fs.existsSync(profileDir)) {
    // 如果已存在，简单清理一下旧的锁文件
    try { fs.rmSync(path.join(profileDir, '.lock'), { force: true }); } catch(e){}
  } else {
    fs.mkdirSync(profileDir, { recursive: true });
  }

  const userDir = path.join(profileDir, 'user');
  const configDir = path.join(userDir, 'config');
  fs.mkdirSync(configDir, { recursive: true });

  // 1. 写入字体替换表 (fontsubs.xml)
  // 这是解决“固定行高”重叠的关键：它告诉 LO 在计算排版宽度时，直接使用 Mac 字体的度量。
  const fontSubsXml = `<?xml version="1.0" encoding="UTF-8"?>
<substitutions>
  <substitution from="Microsoft YaHei"    to="PingFang SC" />
  <substitution from="Microsoft YaHei UI" to="PingFang SC" />
  <substitution from="SimSun"             to="Songti SC" />
  <substitution from="NSimSun"            to="Songti SC" />
  <substitution from="SimHei"             to="Heiti SC" />
  <substitution from="FangSong"           to="STFangsong" />
  <substitution from="KaiTi"              to="Kaiti SC" />
  <substitution from="Microsoft JhengHei" to="PingFang TC" />
  <substitution from="MingLiU"            to="Songti TC" />
</substitutions>`;
  fs.writeFileSync(path.join(configDir, 'fontsubs.xml'), fontSubsXml);

  // 2. 预热 Profile (让 LO 初始化其内部字体数据库)
  try {
    execFileSync(sobinary, [
      '--headless',
      '--norestore',
      '-env:UserInstallation=file://' + profileDir,
      '--version'
    ], { stdio: 'ignore', timeout: 10000 });
  } catch (e) {}

  return profileDir;
}

module.exports = { buildLoProfile };
