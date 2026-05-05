"""
검증기 패키지
모든 검증기를 임포트하고 template 이름을 키로 하는 맵을 제공합니다.
"""

from backend.validators.sum_validator import SumValidator
from backend.validators.outlier_validator import OutlierValidator
from backend.validators.duplicate_validator import DuplicateValidator
from backend.validators.range_validator import RangeValidator
from backend.validators.required_validator import RequiredValidator
from backend.validators.custom_validator import CustomValidator

# template 이름 -> 검증기 클래스 매핑
VALIDATOR_MAP = {
    'sum_check': SumValidator,
    'outlier': OutlierValidator,
    'duplicate': DuplicateValidator,
    'range_check': RangeValidator,
    'required': RequiredValidator,
    'custom': CustomValidator,
}

__all__ = [
    'SumValidator',
    'OutlierValidator',
    'DuplicateValidator',
    'RangeValidator',
    'RequiredValidator',
    'CustomValidator',
    'VALIDATOR_MAP',
]
