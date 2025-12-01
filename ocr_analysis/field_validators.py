"""
Field validators for invoice data patterns.
Validates extracted OCR data against expected formats.
"""

import re
from typing import Tuple


class FieldValidator:
    """Validate extracted invoice fields against required patterns"""
    
    # Pattern definitions based on Excel template analysis
    PATTERNS = {
        'invoice_number': r'^\d{9,10}$',  # 9-10 digit numeric (BELEG_NR)
        'date_yyyymmdd': r'^\d{8}$',   # YYYYMMDD format (BELEG_DAT)
        'debitor_number': r'^\d{9,11}$',  # 9-11 digit numeric (DEBI_KREDI)
        'amount_decimal': r'^\d+(\.\d{2})?$',  # decimal with 2 places
        'cost_center': r'^\d{4}\s+[A-Z]{2}\s+[A-Z]{2}',  # e.g., "1025 TG OB"
        'cost_bearer': r'^\d{12}$',    # 12-digit numeric (KOSTTRAGER)
    }
    
    @staticmethod
    def validate_invoice_number(value: str) -> Tuple[bool, str]:
        """
        Validate invoice/beleg number format.
        
        Args:
            value: Invoice number string
            
        Returns:
            (is_valid, error_message)
        """
        if not value:
            return False, "Invoice number is empty"
        
        # Clean value
        clean = value.strip()
        
        # Check 9-10 digit pattern
        if re.match(FieldValidator.PATTERNS['invoice_number'], clean):
            return True, ""
        
        # Check if numeric but wrong length
        if clean.isdigit():
            return (
                False,
                f"Must be 9-10 digits (got {len(clean)})"
            )
        
        return False, "Must contain only digits"
    
    @staticmethod
    def validate_date(value: str) -> Tuple[bool, str]:
        """
        Validate date in YYYYMMDD format.
        
        Args:
            value: Date string
            
        Returns:
            (is_valid, error_message)
        """
        if not value:
            return False, "Date is empty"
        
        clean = value.strip()
        
        if not re.match(
            FieldValidator.PATTERNS['date_yyyymmdd'],
            clean
        ):
            return False, "Must be YYYYMMDD format (8 digits)"
        
        # Basic range check
        year = int(clean[:4])
        month = int(clean[4:6])
        day = int(clean[6:8])
        
        if not (2000 <= year <= 2099):
            return False, f"Invalid year: {year}"
        if not (1 <= month <= 12):
            return False, f"Invalid month: {month}"
        if not (1 <= day <= 31):
            return False, f"Invalid day: {day}"
        
        return True, ""
    
    @staticmethod
    def validate_debitor(value: str) -> Tuple[bool, str]:
        """
        Validate debitor number format.
        
        Args:
            value: Debitor number string
            
        Returns:
            (is_valid, error_message)
        """
        if not value:
            return False, "Debitor number is empty"
        
        clean = value.strip()
        
        if re.match(FieldValidator.PATTERNS['debitor_number'], clean):
            return True, ""
        
        if clean.isdigit():
            return (
                False,
                f"Must be 9-11 digits (got {len(clean)})"
            )
        
        return False, "Must contain only digits"
    
    @staticmethod
    def validate_amount(value: str) -> Tuple[bool, str]:
        """
        Validate amount format (decimal cents).
        
        Args:
            value: Amount string (numeric, no separators)
            
        Returns:
            (is_valid, error_message)
        """
        if not value:
            return False, "Amount is empty"
        
        clean = str(value).strip()
        
        # Allow integers or decimals with max 2 places
        if re.match(r'^\d+(\.\d{1,2})?$', clean):
            return True, ""
        
        return False, "Must be numeric (e.g., 12345 or 12345.67)"
    
    @staticmethod
    def validate_booking_text(value: str) -> Tuple[bool, str]:
        """
        Validate booking text format.
        Expected: <cost_center> <TG/Service> <Location> <Name>
        
        Args:
            value: Booking text string
            
        Returns:
            (is_valid, error_message)
        """
        if not value:
            return False, "Booking text is empty"
        
        clean = value.strip()
        
        # Check for cost center pattern at start
        if re.match(FieldValidator.PATTERNS['cost_center'], clean):
            return True, ""
        
        # Allow manual entry if pattern not strict
        if len(clean) >= 10:
            return True, ""
        
        return (
            False,
            "Format: <code> TG <location> <name>"
        )
    
    @staticmethod
    def calculate_field_confidence(
        value: str,
        field_type: str,
        ocr_confidence: float = 0.0
    ) -> float:
        """
        Calculate confidence score for extracted field.
        
        Args:
            value: Field value
            field_type: Type of field (invoice_number, date, etc.)
            ocr_confidence: Base OCR confidence (0-100)
            
        Returns:
            Confidence score (0-100)
        """
        if not value:
            return 0.0
        
        # Start with OCR confidence
        score = ocr_confidence
        
        # Adjust based on pattern match
        validators = {
            'invoice_number': FieldValidator.validate_invoice_number,
            'date': FieldValidator.validate_date,
            'debitor': FieldValidator.validate_debitor,
            'amount': FieldValidator.validate_amount,
            'booking_text': FieldValidator.validate_booking_text,
        }
        
        if field_type in validators:
            is_valid, _ = validators[field_type](value)
            if is_valid:
                score = min(100.0, score + 20.0)
            else:
                score = max(0.0, score - 30.0)
        
        return round(score, 1)
    
    @staticmethod
    def validate_all_fields(row_data: dict) -> dict:
        """
        Validate all fields in a row and return validation results.
        
        Args:
            row_data: Dictionary with invoice row data
            
        Returns:
            Dictionary with validation results:
            {
                'field_name': {
                    'valid': bool,
                    'error': str,
                    'confidence': float
                }
            }
        """
        results = {}
        
        # Map field names to validators
        field_map = {
            'BELEG_NR': ('invoice_number', 'invoice_number'),
            'BELEG_DAT': ('date', 'date'),
            'DEBI_KREDI': ('debitor', 'debitor_number'),
            'BETRAG': ('amount', 'amount'),
            'BUCH_TEXT': ('booking_text', 'booking_text'),
        }
        
        ocr_conf = row_data.get('ocr_confidence', 0.0)
        
        for field_name, (validator_key, field_type) in field_map.items():
            value = str(row_data.get(field_name, '')).strip()
            
            # Run validation
            validator_method = getattr(
                FieldValidator,
                f"validate_{validator_key}"
            )
            is_valid, error = validator_method(value)
            
            # Calculate confidence
            confidence = FieldValidator.calculate_field_confidence(
                value,
                validator_key,
                ocr_conf
            )
            
            results[field_name] = {
                'valid': is_valid,
                'error': error,
                'confidence': confidence,
                'value': value
            }
        
        return results
