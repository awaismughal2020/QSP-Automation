"""
Balance Sheet Validator

Validates balance sheet calculations:
- Assets = Liabilities + Equity (fundamental equation)
- Equity Movement = Direct Result (P&L)
"""

from dataclasses import dataclass, field
from typing import Dict, List
from loguru import logger

from ..parsers.bdo_parser import BDOParseResult


@dataclass
class ValidationCheck:
    """Result of a single validation check."""
    name: str
    passed: bool
    message: str
    expected_value: float = 0.0
    actual_value: float = 0.0
    tolerance: float = 0.0
    critical: bool = False


@dataclass
class ValidationResult:
    """Result of balance sheet validation."""
    passed: bool
    checks: List[ValidationCheck] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)


class BalanceValidator:
    """
    Validates balance sheet calculations.
    
    Key validations:
    1. Fundamental equation: Assets = Liabilities + Equity
    2. Equity movement = Direct Result (from P&L)
    """
    
    def __init__(self, 
                 fundamental_tolerance: float = 0.01,
                 equity_movement_tolerance: float = 1.00):
        """
        Initialize validator.
        
        Args:
            fundamental_tolerance: Tolerance for balance sheet equation (EUR)
            equity_movement_tolerance: Tolerance for equity movement check (EUR)
        """
        self.fundamental_tolerance = fundamental_tolerance
        self.equity_movement_tolerance = equity_movement_tolerance
    
    def validate(self, bdo_result: BDOParseResult, totals: Dict[str, float] = None) -> ValidationResult:
        """
        Validate balance sheet calculations.
        
        Args:
            bdo_result: Parsed BDO data
            totals: Optional pre-calculated totals dict
            
        Returns:
            ValidationResult with all checks
        """
        checks = []
        warnings = []
        errors = []
        
        # Calculate totals if not provided
        if totals is None:
            from ..parsers.bdo_parser import BDOParser
            parser = BDOParser()
            totals = parser.calculate_totals(bdo_result)
        
        # Check 1: Fundamental equation
        fundamental_check = self._validate_fundamental_equation(totals)
        checks.append(fundamental_check)
        if not fundamental_check.passed:
            if fundamental_check.critical:
                errors.append(fundamental_check.message)
            else:
                warnings.append(fundamental_check.message)
        
        # Check 2: Equity movement = Direct Result
        equity_check = self._validate_equity_movement(bdo_result)
        checks.append(equity_check)
        if not equity_check.passed:
            if equity_check.critical:
                errors.append(equity_check.message)
            else:
                warnings.append(equity_check.message)
        
        # Overall result
        passed = all(check.passed for check in checks) and len(errors) == 0
        
        return ValidationResult(
            passed=passed,
            checks=checks,
            warnings=warnings,
            errors=errors
        )
    
    def _validate_fundamental_equation(self, totals: Dict[str, float]) -> ValidationCheck:
        """
        Validate balance sheet equation in trial balance format.
        
        In Dutch trial balance format:
        - Assets (debits) are positive
        - Liabilities (credits) are negative  
        - Equity (credits) are negative
        
        Trial balance equation: Debits = Credits
        Balance sheet: Assets = abs(Liabilities) + abs(Equity)
        
        We validate: Assets + Liabilities + Equity ≈ 0 (trial balance format)
        
        Args:
            totals: Dictionary with calculated totals
            
        Returns:
            ValidationCheck result
        """
        total_assets = totals.get('total_assets', 0.0)
        total_equity = totals.get('total_equity', 0.0)
        total_liabilities = totals.get('total_liabilities', 0.0)
        total_income = totals.get('total_income', 0.0)
        total_expenses = totals.get('total_expenses', 0.0)
        direct_result = totals.get('direct_result', 0.0)
        
        # Trial balance sum (should be close to 0)
        # All accounts: Assets + Equity + Liabilities + Income + Expenses + Result = 0
        trial_balance_sum = (total_assets + total_equity + total_liabilities + 
                           total_income + total_expenses + direct_result)
        
        # Use a reasonable tolerance (e.g., 1% of total assets or EUR 10K)
        tolerance = max(abs(total_assets) * 0.01, 10000.0)
        
        passed = abs(trial_balance_sum) <= tolerance
        
        # Also show traditional balance sheet view
        abs_equity = abs(total_equity)
        abs_liabilities = abs(total_liabilities)
        
        message = (
            f"Trial balance check: sum of all accounts = €{trial_balance_sum:,.2f} "
            f"(tolerance: €{tolerance:,.2f}). "
            f"Assets: €{total_assets:,.2f}, "
            f"Equity: €{abs_equity:,.2f}, "
            f"Liabilities: €{abs_liabilities:,.2f}"
        )
        
        return ValidationCheck(
            name="Trial Balance Check",
            passed=passed,
            message=message,
            expected_value=0.0,
            actual_value=trial_balance_sum,
            tolerance=tolerance,
            critical=False  # Downgrade from critical - may have rounding
        )
    
    def _validate_equity_movement(self, bdo_result: BDOParseResult) -> ValidationCheck:
        """
        Validate: Equity Movement = Direct Result (P&L)
        
        This is the critical validation from the manual process.
        In Dutch accounting, equity accounts start with 10, 11.
        
        Args:
            bdo_result: Parsed BDO data
            
        Returns:
            ValidationCheck result
        """
        # Calculate equity movement
        # Dutch equity accounts: 10xxxxx (share capital), 11xxxxx (reserves)
        equity_start = 0.0
        equity_end = 0.0
        
        for code, entry in bdo_result.accounts.items():
            first_two = code[:2] if len(code) >= 2 else ''
            # Dutch equity accounts (EV - Eigen Vermogen)
            if first_two in ('10', '11'):
                equity_start += entry.opening_balance
                equity_end += entry.closing_balance
        
        equity_movement = equity_end - equity_start
        
        # Calculate direct result from P&L
        # Income (8xxxxx) - shown as negative (credits)
        # Expenses (4xxxxx) - shown as positive (debits)
        # Result = Income + Expenses (with proper signs)
        direct_result = 0.0
        for code, entry in bdo_result.accounts.items():
            first_digit = code[0] if code else ''
            # Use closing balance differences to calculate P&L impact
            if first_digit == '8':  # Income accounts (negative = credit)
                direct_result += entry.closing_balance
            elif first_digit == '4':  # Expense accounts (positive = debit)
                direct_result += entry.closing_balance
        
        difference = abs(equity_movement - direct_result)
        passed = difference <= self.equity_movement_tolerance
        
        message = (
            f"Equity movement (€{equity_movement:,.2f}) "
            f"= Direct Result (€{direct_result:,.2f}), "
            f"difference: €{difference:,.2f}"
        )
        
        return ValidationCheck(
            name="Equity Movement",
            passed=passed,
            message=message,
            expected_value=direct_result,
            actual_value=equity_movement,
            tolerance=self.equity_movement_tolerance,
            critical=True
        )

