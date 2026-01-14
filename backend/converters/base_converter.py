"""基础转换器抽象类"""
from abc import ABC, abstractmethod
from typing import Optional
import os


class BaseConverter(ABC):
    """所有转换器的基类"""
    
    @abstractmethod
    def convert(self, input_file: str, output_file: Optional[str] = None) -> str:
        """
        转换文件
        
        Args:
            input_file: 输入文件路径
            output_file: 输出文件路径（可选，默认与输入文件同名不同扩展名）
            
        Returns:
            str: 输出文件路径
        """
        pass
    
    @staticmethod
    def validate_file(file_path: str, extensions: list) -> bool:
        """
        验证文件是否存在且格式正确
        
        Args:
            file_path: 文件路径
            extensions: 允许的扩展名列表
            
        Returns:
            bool: 文件是否有效
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in extensions:
            raise ValueError(f"不支持的文件格式: {ext}, 支持的格式: {extensions}")
        
        return True
    
    @staticmethod
    def get_output_path(input_file: str, output_ext: str, output_file: Optional[str] = None) -> str:
        """
        生成输出文件路径
        
        Args:
            input_file: 输入文件路径
            output_ext: 输出文件扩展名（包含点号，如 '.pdf'）
            output_file: 指定的输出文件路径（可选）
            
        Returns:
            str: 输出文件路径
        """
        if output_file:
            return output_file
        
        base_name = os.path.splitext(input_file)[0]
        return base_name + output_ext
