from django.contrib import admin
from .models import Instrument, Portfolio, CashEquivalents

class InstrumentAdmin(admin.ModelAdmin):
    list_display = ('instrument_name', 'industry_rating', 'quantity', 'market_value', 'isin')
    search_fields = ('instrument_name', 'isin')
    
class PortfolioAdmin(admin.ModelAdmin):
    list_display = ('name', 'total_market_value', 'equity_total', 'debt_total', 'other_total', 'portfolio_date')
    list_filter = ('portfolio_date',)
    
class CashEquivalentsAdmin(admin.ModelAdmin):
    list_display = ('total_value', 'portfolio')

# Register the models with their customized admin classes
admin.site.register(Instrument, InstrumentAdmin)
admin.site.register(Portfolio, PortfolioAdmin)
admin.site.register(CashEquivalents, CashEquivalentsAdmin)
