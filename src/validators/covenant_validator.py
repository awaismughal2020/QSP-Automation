"""
Covenant Validator

Validates covenant compliance against thresholds.
Checks for warnings when values are within 5% of threshold.
"""

from dataclasses import dataclass, field
from typing import List, Dict
from loguru import logger

from ..transformers.compliance_calc import ComplianceResult, CovenantResult, CovenantStatus


@dataclass
class CovenantValidationIssue:
    """Represents a covenant validation issue."""
    covenant_name: str
    covenant_id: str
    value: float
    threshold: float
    status: CovenantStatus
    message: str


@dataclass
class CovenantValidationResult:
    """Result of covenant validation."""
    passed: bool
    failed_covenants: List[CovenantValidationIssue] = field(default_factory=list)
    warning_covenants: List[CovenantValidationIssue] = field(default_factory=list)
    all_passed: List[str] = field(default_factory=list)


class CovenantValidator:
    """
    Validates covenant compliance results.
    
    Checks:
    - All covenants pass thresholds
    - Warnings for covenants within 5% of threshold
    """
    
    def __init__(self, warning_margin: float = 0.05):
        """
        Initialize validator.
        
        Args:
            warning_margin: Percentage margin for warning (default 5%)
        """
        self.warning_margin = warning_margin
    
    def validate(self, compliance_result: ComplianceResult) -> CovenantValidationResult:
        """
        Validate covenant compliance results.
        
        Args:
            compliance_result: ComplianceResult from ComplianceCalculator
            
        Returns:
            CovenantValidationResult with failed and warning covenants
        """
        failed_covenants = []
        warning_covenants = []
        all_passed = []
        
        for covenant_id, covenant in compliance_result.covenants.items():
            if covenant.status == CovenantStatus.FAIL:
                issue = CovenantValidationIssue(
                    covenant_name=covenant.name,
                    covenant_id=covenant_id,
                    value=covenant.value,
                    threshold=covenant.threshold,
                    status=covenant.status,
                    message=self._format_failure_message(covenant)
                )
                failed_covenants.append(issue)
            elif covenant.status == CovenantStatus.WARNING:
                issue = CovenantValidationIssue(
                    covenant_name=covenant.name,
                    covenant_id=covenant_id,
                    value=covenant.value,
                    threshold=covenant.threshold,
                    status=covenant.status,
                    message=self._format_warning_message(covenant)
                )
                warning_covenants.append(issue)
            else:
                all_passed.append(covenant.name)
        
        passed = len(failed_covenants) == 0
        
        return CovenantValidationResult(
            passed=passed,
            failed_covenants=failed_covenants,
            warning_covenants=warning_covenants,
            all_passed=all_passed
        )
    
    def _format_failure_message(self, covenant: CovenantResult) -> str:
        """Format failure message for a covenant."""
        if covenant.operator == '>=':
            return (
                f"{covenant.name} FAILED: {covenant.value:.2%} < {covenant.threshold:.0%} threshold. "
                f"Shortfall: {(covenant.threshold - covenant.value) / covenant.threshold:.1%}"
            )
        else:  # '<='
            return (
                f"{covenant.name} FAILED: {covenant.value:.1%} > {covenant.threshold:.0%} threshold. "
                f"Excess: {(covenant.value - covenant.threshold) / covenant.threshold:.1%}"
            )
    
    def _format_warning_message(self, covenant: CovenantResult) -> str:
        """Format warning message for a covenant."""
        margin_pct = abs(covenant.margin) * 100
        return (
            f"{covenant.name} WARNING: {covenant.value:.2%} is within {margin_pct:.1f}% "
            f"of threshold {covenant.threshold:.0%}. Monitor closely."
        )
    
    def get_summary(self, validation_result: CovenantValidationResult) -> str:
        """
        Get human-readable summary of validation results.
        
        Args:
            validation_result: CovenantValidationResult
            
        Returns:
            Formatted summary string
        """
        lines = []
        
        if validation_result.passed:
            lines.append("✓ All covenants PASSED")
        else:
            lines.append(f"✗ {len(validation_result.failed_covenants)} covenant(s) FAILED")
        
        if validation_result.warning_covenants:
            lines.append(f"⚠ {len(validation_result.warning_covenants)} covenant(s) with WARNINGS")
        
        if validation_result.failed_covenants:
            lines.append("\nFailed Covenants:")
            for issue in validation_result.failed_covenants:
                lines.append(f"  • {issue.message}")
        
        if validation_result.warning_covenants:
            lines.append("\nWarning Covenants:")
            for issue in validation_result.warning_covenants:
                lines.append(f"  • {issue.message}")
        
        if validation_result.all_passed:
            lines.append(f"\nPassed Covenants ({len(validation_result.all_passed)}):")
            for name in validation_result.all_passed:
                lines.append(f"  ✓ {name}")
        
        return "\n".join(lines)

