from django import template
register = template.Library()
@register.filter
def uppercase(value):
    return value.upper()
@register.filter
def _slice(value, slice_range):
    return value[slice_range]
