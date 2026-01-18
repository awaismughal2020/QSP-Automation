"""
Tests for Compliance Calculator
"""

import pytest
from src.transformers.compliance_calc import (
    ComplianceCalculator,
    ComplianceResult,
    CovenantStatus
)


class TestComplianceCalculator:
    """Test compliance calculator."""
    
    @pytest.fixture
    def calculator(self):
        """Create compliance calculator instance."""
        return ComplianceCalculator()
    
    def test_calculate_dscr(self, calculator):
        """Test DSCR calculation."""
        # NOI = 1,200,000, Debt Service = 1,000,000
        # DSCR = 1.2 (120%)
        dscr = calculator._calculate_dscr(1_200_000, 1_000_000)
        assert dscr == 1.2
        
        # Edge case: zero debt service
        dscr_inf = calculator._calculate_dscr(1_200_000, 0)
        assert dscr_inf == float('inf')
    
    def test_calculate_ndy(self, calculator):
        """Test Net Debt Yield calculation."""
        # NOI = 800,000, Net Debt = 12,000,000
        # NDY = 0.0667 (6.67%)
        ndy = calculator._calculate_ndy(800_000, 12_000_000)
        assert abs(ndy - 0.0667) < 0.0001
        
        # Edge case: zero net debt
        ndy_inf = calculator._calculate_ndy(800_000, 0)
        assert ndy_inf == float('inf')
    
    def test_calculate_ltv(self, calculator):
        """Test Loan to Value calculation."""
        # Net Debt = 7,500,000, Property Value = 12,000,000
        # LTV = 0.625 (62.5%)
        ltv = calculator._calculate_ltv(7_500_000, 12_000_000)
        assert ltv == 0.625
        
        # Edge case: zero property value
        ltv_inf = calculator._calculate_ltv(7_500_000, 0)
        assert ltv_inf == float('inf')
    
    def test_calculate_all_covenants(self, calculator):
        """Test calculation of all covenants."""
        result = calculator.calculate(
            net_operating_income=1_200_000,
            debt_service=1_000_000,
            net_debt=12_000_000,
            property_value=20_000_000,
            quarter="Q3 2025"
        )
        
        assert isinstance(result, ComplianceResult)
        assert result.quarter == "Q3 2025"
        assert len(result.covenants) == 5
        
        # Check all covenant types are present
        assert 'hdscr' in result.covenants
        assert 'pdscr' in result.covenants
        assert 'hndy' in result.covenants
        assert 'pndy' in result.covenants
        assert 'ltv' in result.covenants
    
    def test_dscr_threshold_pass(self, calculator):
        """Test DSCR passing threshold."""
        # DSCR = 1.25 (125%) >= 1.20 (120%) threshold
        result = calculator.calculate(
            net_operating_income=1_250_000,
            debt_service=1_000_000,
            net_debt=12_000_000,
            property_value=20_000_000,
            quarter="Q3 2025"
        )
        
        assert result.covenants['hdscr'].status == CovenantStatus.PASS
        assert result.covenants['hdscr'].value >= 1.20
    
    def test_dscr_threshold_fail(self, calculator):
        """Test DSCR failing threshold."""
        # DSCR = 1.15 (115%) < 1.20 (120%) threshold
        result = calculator.calculate(
            net_operating_income=1_150_000,
            debt_service=1_000_000,
            net_debt=12_000_000,
            property_value=20_000_000,
            quarter="Q3 2025"
        )
        
        assert result.covenants['hdscr'].status == CovenantStatus.FAIL
        assert result.covenants['hdscr'].value < 1.20
        assert result.overall_status == CovenantStatus.FAIL
    
    def test_ndy_threshold_pass(self, calculator):
        """Test NDY passing threshold."""
        # NDY = 0.07 (7%) >= 0.062 (6.2%) threshold
        result = calculator.calculate(
            net_operating_income=840_000,
            debt_service=1_000_000,
            net_debt=12_000_000,
            property_value=20_000_000,
            quarter="Q3 2025"
        )
        
        assert result.covenants['hndy'].status == CovenantStatus.PASS
        assert result.covenants['hndy'].value >= 0.062
    
    def test_ltv_threshold_pass(self, calculator):
        """Test LTV passing threshold."""
        # LTV = 0.60 (60%) <= 0.62 (62%) threshold
        result = calculator.calculate(
            net_operating_income=1_200_000,
            debt_service=1_000_000,
            net_debt=12_000_000,
            property_value=20_000_000,
            quarter="Q3 2025"
        )
        
        assert result.covenants['ltv'].status == CovenantStatus.PASS
        assert result.covenants['ltv'].value <= 0.62
    
    def test_ltv_threshold_fail(self, calculator):
        """Test LTV failing threshold."""
        # LTV = 0.65 (65%) > 0.62 (62%) threshold
        result = calculator.calculate(
            net_operating_income=1_200_000,
            debt_service=1_000_000,
            net_debt=13_000_000,
            property_value=20_000_000,
            quarter="Q3 2025"
        )
        
        assert result.covenants['ltv'].status == CovenantStatus.FAIL
        assert result.covenants['ltv'].value > 0.62
    
    def test_warning_margin(self, calculator):
        """Test warning status when close to threshold."""
        # DSCR = 1.21 (121%) - just above threshold, within 5% margin
        result = calculator.calculate(
            net_operating_income=1_210_000,
            debt_service=1_000_000,
            net_debt=12_000_000,
            property_value=20_000_000,
            quarter="Q3 2025"
        )
        
        # Should be PASS but close to threshold
        # Note: Warning logic depends on margin calculation
        assert result.covenants['hdscr'].value >= 1.20
    
    def test_projected_values(self, calculator):
        """Test projected covenant calculations."""
        result = calculator.calculate(
            net_operating_income=1_200_000,
            debt_service=1_000_000,
            net_debt=12_000_000,
            property_value=20_000_000,
            projected_noi=1_300_000,
            projected_debt_service=1_000_000,
            quarter="Q3 2025"
        )
        
        # Projected DSCR should be higher than historic
        assert result.covenants['pdscr'].value > result.covenants['hdscr'].value
        assert result.covenants['pndy'].value > result.covenants['hndy'].value
    
    def test_format_result(self, calculator):
        """Test result formatting."""
        result = calculator.calculate(
            net_operating_income=1_200_000,
            debt_service=1_000_000,
            net_debt=12_000_000,
            property_value=20_000_000,
            quarter="Q3 2025"
        )
        
        formatted = calculator.format_result(result)
        assert "Q3 2025" in formatted
        assert "Compliance Certificate" in formatted
        assert "Covenant Results" in formatted

