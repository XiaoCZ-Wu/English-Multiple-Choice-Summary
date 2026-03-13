"""
配置数据模型
管理应用程序配置
"""

import json
import os
from dataclasses import dataclass, field, asdict
from typing import Dict, Any, List

from src.utils import DEFAULT_CONFIG, CONFIG_FILE, app_logger

# 当前配置文件版本（与软件版本保持一致）
CONFIG_VERSION = "1.3"


@dataclass
class AIConfig:
    """AI配置类"""
    name: str = ""  # AI名称
    base_url: str = ""  # Base URL
    api_key: str = ""  # API Key
    model: str = ""  # 模型名称

    def to_dict(self) -> dict:
        """转换为字典 - 顺序：name、base_url、model、api_key"""
        return {
            "name": self.name,
            "base_url": self.base_url,
            "model": self.model,
            "api_key": self.api_key
        }

    @classmethod
    def from_dict(cls, data: dict) -> "AIConfig":
        """从字典创建"""
        return cls(
            name=data.get("name", ""),
            base_url=data.get("base_url", ""),
            api_key=data.get("api_key", ""),
            model=data.get("model", "")
        )


@dataclass
class AppConfig:
    """应用配置类"""
    version: str = CONFIG_VERSION  # 配置文件版本
    font_name: str = "Microsoft YaHei"
    font_size: int = 11
    output_dir: str = ".\\output\\"
    ai_configs: List[AIConfig] = field(default_factory=list)  # AI配置列表
    ocr_ai_name: str = ""  # OCR使用的AI名称
    chat_ai_name: str = ""  # 提问使用的AI名称
    lan_port: int = 8080  # 局域网通信端口

    def __post_init__(self):
        """确保输出目录以分隔符结尾"""
        if self.output_dir and not (self.output_dir.endswith("/") or self.output_dir.endswith("\\")):
            self.output_dir += "\\"

    @classmethod
    def from_dict(cls, data: dict) -> "AppConfig":
        """从字典创建配置，支持自动升级"""
        # 检查版本，如有需要则升级
        config_version = data.get("version", "0.0")
        if config_version != CONFIG_VERSION:
            app_logger.info(f"[Config] 检测到旧版本配置文件 ({config_version})，正在升级到 {CONFIG_VERSION}")
            app_logger.info(f"[Config] 原配置字段: {list(data.keys())}")
            data = cls._upgrade_config(data, config_version)
            app_logger.info(f"[Config] 升级后配置字段: {list(data.keys())}")

        # 解析AI配置列表
        ai_configs_data = data.get("ai_configs", [])
        ai_configs = [AIConfig.from_dict(cfg) for cfg in ai_configs_data]

        # 处理 output_dir（支持 output_dir 和 output-dir 两种键名）
        output_dir = data.get("output_dir")
        if output_dir is None:
            output_dir = data.get("output-dir", DEFAULT_CONFIG.get("output-dir", ".\\output\\"))

        return cls(
            version=data.get("version", CONFIG_VERSION),
            font_name=data.get("font-name", DEFAULT_CONFIG["font-name"]),
            font_size=data.get("font-size", DEFAULT_CONFIG["font-size"]),
            output_dir=output_dir,
            ai_configs=ai_configs,
            ocr_ai_name=data.get("ocr_ai_name", ""),
            chat_ai_name=data.get("chat_ai_name", ""),
            lan_port=data.get("lan_port", 8080)
        )

    @staticmethod
    def _upgrade_config(data: dict, old_version: str) -> dict:
        """升级配置文件到最新版本"""
        # 从默认配置开始
        upgraded = dict(DEFAULT_CONFIG)
        upgraded["version"] = CONFIG_VERSION

        # 复制原配置中所有存在的字段（保留用户数据）
        for key in data:
            if key != "version":  # 版本号使用新的
                upgraded[key] = data[key]

        # 处理 ai_configs 中的每个配置（确保有新字段）
        if "ai_configs" in upgraded and upgraded["ai_configs"]:
            for cfg in upgraded["ai_configs"]:
                # 确保每个AI配置都有 model 字段
                if "model" not in cfg:
                    cfg["model"] = ""

        app_logger.info(f"[Config] 配置文件升级完成: {old_version} -> {CONFIG_VERSION}")
        return upgraded

    def to_dict(self) -> dict:
        """转换为字典（保持与原JSON格式兼容）"""
        return {
            "version": self.version,
            "font-name": self.font_name,
            "font-size": self.font_size,
            "output_dir": self.output_dir,
            "ai_configs": [cfg.to_dict() for cfg in self.ai_configs],
            "ocr_ai_name": self.ocr_ai_name,
            "chat_ai_name": self.chat_ai_name,
            "lan_port": self.lan_port
        }

    def get_ai_config(self, name: str) -> AIConfig:
        """根据名称获取AI配置"""
        for cfg in self.ai_configs:
            if cfg.name == name:
                return cfg
        return AIConfig()

    def add_ai_config(self, name: str, base_url: str, api_key: str) -> bool:
        """添加AI配置"""
        # 检查名称是否已存在
        for cfg in self.ai_configs:
            if cfg.name == name:
                return False
        self.ai_configs.append(AIConfig(name=name, base_url=base_url, api_key=api_key))
        return True

    def update_ai_config(self, old_name: str, name: str, base_url: str, api_key: str) -> bool:
        """更新AI配置"""
        for cfg in self.ai_configs:
            if cfg.name == old_name:
                cfg.name = name
                cfg.base_url = base_url
                cfg.api_key = api_key
                # 更新引用
                if self.ocr_ai_name == old_name:
                    self.ocr_ai_name = name
                if self.chat_ai_name == old_name:
                    self.chat_ai_name = name
                return True
        return False

    def delete_ai_config(self, name: str) -> bool:
        """删除AI配置"""
        for i, cfg in enumerate(self.ai_configs):
            if cfg.name == name:
                del self.ai_configs[i]
                # 清除引用
                if self.ocr_ai_name == name:
                    self.ocr_ai_name = ""
                if self.chat_ai_name == name:
                    self.chat_ai_name = ""
                return True
        return False

    @classmethod
    def load(cls, filepath: str = CONFIG_FILE) -> "AppConfig":
        """从文件加载配置"""
        app_logger.info(f"[Config] 尝试加载配置: {filepath}")
        try:
            if os.path.exists(filepath):
                app_logger.info(f"[Config] 配置文件存在，正在读取...")
                with open(filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                app_logger.info(f"[Config] 配置加载成功: {data.get('version', 'unknown')}")
                return cls.from_dict(data)
            else:
                app_logger.info(f"[Config] 配置文件不存在: {filepath}")
        except Exception as e:
            app_logger.error(f"[Config] 加载配置失败: {e}")
        return cls()  # 返回默认配置

    def save(self, filepath: str = CONFIG_FILE):
        """保存配置到文件"""
        try:
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            app_logger.error(f"保存配置失败: {e}")
            return False
