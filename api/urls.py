from api.views import index, getBook, userType, filterByStalls, filterBooks, recommendApi, purchaseApi, getRecommendation, getOrders, import_book_data, getAdminCount, bookAction, booksActionsSeller, recommendExcelApiAll, recommendExcelApiSeller, recommendExcelApiSellerAndSubject, recommendExcelApiSubject
from django.urls import path

urlpatterns = [
    path('', index, name = 'index'),
    path('user/type/', userType, name = 'userType'),
    path('book/<int:book_id>/', getBook, name='getBook'),
    path('stalls/', filterByStalls, name = 'filterStalls'),
    path('books/', filterBooks, name = 'filterBooks' ),

    path( 'recommend/', recommendApi, name = 'Recommend' ),
    path( 'purchase/', purchaseApi, name = 'Purchase Book' ),

    path( 'recommendations/', getRecommendation, name = 'Recommendation' ),
    path( 'orders/', getOrders, name = 'Get Orders' ),

    path('admin/books/import/', import_book_data, name = 'Import Books'),

    path('admin/count/<int:pk>/', getAdminCount, name = "getAdminCount"),
    path('admin/books/action/<int:pk>/', bookAction, name = 'booksActions'),

    path('admin/books/action/seller/<int:sellerid>/<int:type>/',booksActionsSeller, name = 'booksActions'),
    path('admin/<int:type>/excel/all', recommendExcelApiAll, name = 'RecommendApi'),
    path('admin/<int:type>/excel/subject/<str:subject>', recommendExcelApiSubject, name = 'RecommendApi'),
    path('admin/<int:type>/excel/seller/<int:seller>', recommendExcelApiSeller, name = 'RecommendApi'),
    path('admin/<int:type>/excel/<int:seller>/<str:subject>', recommendExcelApiSellerAndSubject, name = 'recommendExcelApiSellerAndSubject'),
]
