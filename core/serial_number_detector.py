"""序号检测器 - 识别2/3/4位序号"""
from typing import List, NamedTuple
import re


class RowInfo(NamedTuple):
    row_index: int
    value: int
    format_code: str
    display_digits: int


class SerialNumberDetector:
    """检测Excel中的序号行"""
    
    def detect(self, ws) -> List[RowInfo]:
        """
        检测工作表中的序号行
        
        Args:
            ws: openpyxl工作表对象
            
        Returns:
            List[RowInfo]: 检测到的序号行信息列表
        """
        results = []
        
        for row_idx in range(1, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=1)
            
            if cell.value != 1:
                continue
            
            num_format = cell.number_format or ''
            
            # 检查是否为数字格式 (00, 000, 0000等)
            if self._is_serial_format(num_format):
                digits = len(num_format)
                results.append(RowInfo(
                    row_index=row_idx,
                    value=cell.value,
                    format_code=num_format,
                    display_digits=digits
                ))
        
        return results
    
    def _is_serial_format(self, num_format: str) -> bool:
        """检查是否为序号格式"""
        # 匹配纯0的格式字符串
        return bool(re.match(r'^0+$', num_format))
