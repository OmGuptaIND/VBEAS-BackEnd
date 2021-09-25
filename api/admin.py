from api.models import Book, ExcelFileUpload, Order, Recommend, Seller, UserType
from django.contrib import admin

# Register your models here.

admin.site.register(Book)
admin.site.register(Seller)
admin.site.register(UserType)
admin.site.register(ExcelFileUpload)
@admin.register(Recommend)
class RecommendAdmin(admin.ModelAdmin):
    list_display = ( 'name', 'email', 'usertype', 'book', 'seller', 'recommended_to_library', 'quantity')

@admin.register(Order)
class OrderAdmin(admin.ModelAdmin):
    list_display = ( 'name', 'email', 'usertype', 'book', 'seller', 'is_ordered', 'quantity')

