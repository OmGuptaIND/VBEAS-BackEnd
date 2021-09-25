import os, re, random, math
from django import template
from django.db.models import query
from django.db.models.query_utils import Q
from django.shortcuts import render
from django.http.response import HttpResponse, JsonResponse
from django.shortcuts import render, redirect
from django.views.decorators.csrf import csrf_exempt
from django.core.mail import send_mail
from django.conf import settings

# Rest Frameworks
from rest_framework.decorators import api_view, permission_classes
from rest_framework.permissions import IsAuthenticated, AllowAny
from rest_framework.response import Response
from rest_framework import status
from django.core.paginator import Paginator

from openpyxl import * 
from openpyxl.writer.excel import save_virtual_workbook
# import pandas as pd
from openpyxl import load_workbook

# Models
from api.models import Book, Seller, Recommend, ExcelFileUpload, UserType, Order

# Create your views here.

def home_index(request):
    return redirect('api/')

def index(request):
    return HttpResponse("Backend Connected")

@api_view(['POST'])
@permission_classes([AllowAny,])
@csrf_exempt
def filterByStalls(request):
    valid_query_parameters = [ 'qid', 'page_number', 'filter_by_subject', 'medium', 'sort_by', 'order_by' ]
    valid_sort_paramters = ['price', 'year_of_publication']

    if request.method == 'POST':
        query = dict( request.data )
        query_parameters = list( query.keys() )
        print(query)

        if 'qid' not in query_parameters or 'page_number' not in query_parameters:
            return Response({
                'status' : 'error',
                'code' : '400',
                'message' : 'qid or page number missing'
            }, status = status.HTTP_200_OK)
        
        for key in query_parameters:
            if key not in valid_query_parameters:
                return Response({
                    'status' : 'error', 
                    'code' : '404',
                    'message' : f'Query is not valid: {key}',
                }, status = status.HTTP_200_OK)
        
        if 'sort_by' in query_parameters and query['sort_by'] not in valid_sort_paramters:
            return Response({
                    'status' : 'error', 
                    'code' : '404',
                    'message' : f"Sort By is not valid: {query['sort_by']}",
                }, status = status.HTTP_200_OK)

        if 'sort_by' in query_parameters and 'order_by' not in query_parameters:
            return Response({
                    'status' : 'error', 
                    'code' : '404',
                    'message' : f"Order By is to passed",
                }, status = status.HTTP_200_OK)

        qid = int(query['qid'])
        page_number = int(query['page_number'])

        if qid not in range(1, 10):
            return Response({
                    'status' : 'error', 
                    'code' : '404',
                    'message' : f"No such seller with this ID exists",
                }, status = status.HTTP_200_OK)

        if 'filter_by_subject' in query_parameters and 'medium' in query_parameters:
            books = Book.objects.filter( seller_id = query['qid'], subject__icontains = query['filter_by_subject'], medium__icontains = query['medium'] ).order_by('id')
        elif 'medium' in query_parameters:
            books = Book.objects.filter( seller_id = query['qid'], medium__icontains = query['medium'] ).order_by('id')
        elif 'filter_by_subject' in query_parameters:
            books = Book.objects.filter( seller_id = query['qid'], subject__icontains = query['filter_by_subject'] ).order_by('id')
        else:
            books = Book.objects.filter(seller_id = query['qid'] ).order_by('id')
        
        if 'sort_by' in query_parameters and query['order_by'] == 'desc':
            books = books.order_by( '-' + str(query['sort_by']))
        elif 'sort_by' in query_parameters:
            books = books.order_by(query['sort_by'])

        listBooks = Paginator(books.values(), 8)
        count = listBooks.count

        if count == 0:
            return Response({
                    'status':"success",
                    'code':200,
                    'message':"Nothing found",
                }, status = status.HTTP_200_OK)
        
        total_page = math.ceil( count / 8 )

        if page_number > total_page:
            return Response({
                'status':"error",
                'code':404,
                'message':"Page Number Exceeded permissible range",
            }, status = status.HTTP_200_OK)

        page = listBooks.page(page_number)

        context = {
            'status':'success',
            'code':200,
            'count': count, 
            'total_page':math.ceil( count / 8 ),
            'current_page':query['page_number'],
            'total_data': len(list(page)) ,
            'data':list(page),
        }

        return Response(context, status = status.HTTP_200_OK)

@api_view(['POST'])
@permission_classes([AllowAny,])
@csrf_exempt
def filterBooks(request):
    valid_query_parameters = [ 'q', 'page_number', 'filter_by_subject', 'medium', 'sort_by', 'order_by' ]
    valid_sort_paramters = ['price', 'year_of_publication']

    if request.method == 'POST':
        query = dict( request.data )
        query_parameters = list( query.keys() )
        print(query)

        if 'q' not in query_parameters or 'page_number' not in query_parameters:
            return Response({
                'status' : 'error',
                'code' : '400',
                'message' : 'q or page number missing'
            }, status = status.HTTP_200_OK)
        
        for key in query_parameters:
            if key not in valid_query_parameters:
                return Response({
                    'status' : 'error', 
                    'code' : '404',
                    'message' : f'Query is not valid: {key}',
                }, status = status.HTTP_200_OK)
        
        if 'sort_by' in query_parameters and query['sort_by'] not in valid_sort_paramters:
            return Response({
                    'status' : 'error', 
                    'code' : '404',
                    'message' : f"Sort By is not valid: {query['sort_by']}",
                }, status = status.HTTP_200_OK)

        if 'sort_by' in query_parameters and 'order_by' not in query_parameters:
            return Response({
                    'status' : 'error', 
                    'code' : '404',
                    'message' : f"Order By is to passed",
                }, status = status.HTTP_200_OK)

        q = query['q']
        page_number = int(query['page_number'])

        try:
            books = Book.objects.filter(subject__icontains = q) | Book.objects.filter( author__icontains = q ) | Book.objects.filter( title__icontains = q )
        
        except:
                return Response({
                    'status':"error",
                    'code':404,
                    'message': 'Something went wrong',
                }, status = status.HTTP_200_OK)

        if 'filter_by_subject' in query_parameters and 'medium' in query_parameters:
            books = books.filter( subject__icontains = query['filter_by_subject'], medium__icontains = query['medium'] ).order_by('id')
        elif 'medium' in query_parameters:
            books = books.filter( medium__icontains = query['medium'] ).order_by('id')
        elif 'filter_by_subject' in query_parameters:
            books = books.filter( subject__icontains = query['filter_by_subject'] ).order_by('id')
        
        if 'sort_by' in query_parameters and query['order_by'] == 'desc':
            books = books.order_by( '-' + str(query['sort_by']))
        elif 'sort_by' in query_parameters:
            books = books.order_by(query['sort_by'])
        
        listBooks = Paginator(books.values(), 8)
        count = listBooks.count

        if count == 0:
            return Response({
                    'status':"success",
                    'code':200,
                    'message':"Nothing found",
                }, status = status.HTTP_200_OK)
        
        total_page = math.ceil( count / 8 )

        if page_number > total_page:
            return Response({
                'status':"error",
                'code':404,
                'message':"Page Number Exceeded permissible range",
            }, status = status.HTTP_200_OK)

        page = listBooks.page(page_number)

        context = {
            'status':'success',
            'code':200,
            'count': count, 
            'total_page':math.ceil( count / 8 ),
            'current_page':query['page_number'],
            'total_data': len(list(page)) ,
            'data':list(page),
        }

        return Response(context, status = status.HTTP_200_OK)

@api_view(['GET'])
@permission_classes((AllowAny,))
def getBook(request, book_id):
    if book_id is None:
        return JsonResponse({
            "status" : "error",
            "code" : 404,
            "message" : "No Such Book Available"
        })
    try:
        book = Book.objects.filter(id=book_id)
        if len(book.values()) == 0:
            return Response({
                "error":"404 NOT Found",
                'message':f"Book Id does not exit ${book_id}"
            }, status = status.HTTP_404_NOT_FOUND)
    except:
        return Response({
            'error':"404 NOT FOUND",
            'message': "Book Id "+str(book_id) + " does not exist"
        }, status = status.HTTP_404_NOT_FOUND)
    return Response({'data': book.values()})


@api_view(['POST'])
@permission_classes((AllowAny,))
def userType(request):
    if request.method == 'POST':
        admins = ['20ume034', 'pahmad', 'shweta.pandey', 'librarian']
        data = request.data
        regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        if re.fullmatch(regex, data['email']):
            email = data['email']
            username, domain = email.split('@')
            if domain == 'lnmiit.ac.in':
                if username in admins:
                    return Response({
                        "status" : 'Success',
                        "code" : 200,
                        "Message" : 'User is admin', 
                        'data' : {
                            'type':'admin',
                            'code' : 1
                        }
                    }, status = status.HTTP_200_OK)
                
                elif username.isalnum() and username[2] in ['p', 'm']:
                    return Response({
                        'status' : 'Success',
                        'code' : 200,
                        'message' : 'User is Research Scholar',
                        'data' : {
                            'type' : 'Scholar', 
                            'code' : 3
                        }
                    }, status=status.HTTP_200_OK)

                elif not username.isalnum():
                    return JsonResponse({
                        "status" : 'Success',
                        "code" : 200,
                        "Message" : 'User is Professor', 
                        'data' : {
                            'type':'Professor',
                            'code' : 2
                        }
                    })
            else:
                return Response({
                    "status" : 'Error',
                    'code': 403,
                    'Message' : " Enter a Valid LNMIIT Email only"
                }, status=status.HTTP_403_FORBIDDEN)
        else:
            return Response({
                "status" : 'Error',
                'code': 403,
                'Message' : " Enter a Valid Email "
            }, status=status.HTTP_403_FORBIDDEN)

        return Response({
            "status" : 'Success',
            "code" : 200,
            "Message" : 'User is Student', 
            'data' : {
                'type':'Student',
                'code' : 4
            }
        }, status=status.HTTP_202_ACCEPTED)


from django.template.loader import render_to_string, get_template
from django.core.mail import EmailMessage

@api_view(['POST'])
@permission_classes((AllowAny,))
@csrf_exempt
def recommendApi(request):
    valid_keys = ['items', 'email', 'name', 'type', 'code']
    query = dict( request.data )
    query_parameters = list( query.keys() )

    for key in valid_keys:
        if key not in query_parameters:
            return Response({
                'status' : "Error",
                'code' : 400,
                'message' : f"Wrong Query Passed {key}"
            }, status=status.HTTP_206_PARTIAL_CONTENT)
    
    regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    email = query['email']
    if re.fullmatch(regex, email):
        username, domain = email.split('@')
        if domain != 'lnmiit.ac.in':
            return Response({
                'status' : "Error",
                'code' : 206,
                'message' : f"Enter a Valid LNMIIT Email Only"
            }, status = status.HTTP_206_PARTIAL_CONTENT)

        items = query['items']
        name = query['name']
        code = query['code']
        books = []
        userType = UserType.objects.get( code = code )
        row = 1
        for item in items:

            book = Book.objects.get( id = item['id'] )
            books.append( { 'id' : row, 'title' : book.title, 'author' : book.author, 'subject' : book.subject, 'publisher' : book.publisher, 'medium' : book.medium, 'demo' : book.price_denomination, 'price' : book. expected_price, 'quantity' : item['qty'] } )
            recommend = Recommend()
            recommend.book = book
            recommend.quantity = item['qty']
            recommend.name = name
            recommend.email = email
            recommend.usertype = userType
            recommend.recommended_to_library = True
            row += 1
            recommend.save()
        ctx = {
            'name' : name,
            'total' : len(books),
            'books' : books
        }
        template = get_template( 'email.html').render(ctx)

        msg = EmailMessage(
            'Thanks For Recommending Books to the Library',
            template,
            settings.EMAIL_HOST_USER,
            [email]
        )
        
        msg.content_subtype = "html" 
        msg.send()
        return Response({
            'status' : "Success",
            'code' : 201,
            'message' : "Added Recommendations",
        }, status = status.HTTP_201_CREATED)

    else:
        return Response({
            'status' : "Error",
            'code' : 206,
            'message' : f"Enter a Valid Email"
        }, status=status.HTTP_206_PARTIAL_CONTENT)


@api_view(['POST'])
@permission_classes((AllowAny,))
@csrf_exempt
def purchaseApi(request):
    valid_keys = ['email', 'name', 'type', 'code', 'book_id', 'qty']
    query = dict( request.data )
    query_parameters = list( query.keys() )

    for key in valid_keys:
        if key not in query_parameters:
            return Response({
                'status' : "Error",
                'code' : 400,
                'message' : f"Wrong Query Passed {key}"
            }, status=status.HTTP_206_PARTIAL_CONTENT)
    
    regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    email = query['email']

    if re.fullmatch(regex, email):
        username, domain = email.split('@')
        if domain != 'lnmiit.ac.in':
            return Response({
                'status' : "Error",
                'code' : 206,
                'message' : f"Enter a Valid LNMIIT Email Only"
            }, status = status.HTTP_206_PARTIAL_CONTENT)
        name = query['name']
        code = query['code']
        book_id = query['book_id']
        qty = query['qty']
        userType = UserType.objects.get( code = code )
        try:
            book = Book.objects.get( id = book_id )
            order = Order()
            order.email = email
            order.name = name
            order.book = book
            order.quantity = qty
            order.is_ordered = True
            order.usertype = userType

            order.save()

            print(f'Book Ordered with ID: {order.id}-{order.quantity}')
            
            return Response({
                'status' : "Success",
                'code' : 201,
                'message' : "Book Ordered Success",
            }, status = status.HTTP_201_CREATED)
        except:
            return Response({
                'status' : 'Error', 
                'code' : 400,
                'message' : "Something Wrong with order table"
            }, status = status.HTTP_400_BAD_REQUEST)



    else:
        return Response({
            'status' : "Error",
            'code' : 206,
            'message' : f"Enter a Valid Email"
        }, status=status.HTTP_206_PARTIAL_CONTENT)

    return Response({
        'status' : 'Success',
        'code' : 201,
        'Message' : "Recieved",
    }, status = status.HTTP_201_CREATED)


@api_view(['POST'])
@permission_classes((AllowAny,))
@csrf_exempt
def getRecommendation(request):
    valid_keys = ['email', 'name']
    query = dict( request.data )
    query_parameters = list( query.keys() )

    for key in valid_keys:
        if key not in query_parameters:
            return Response({
                'status' : "Error",
                'code' : 400,
                'message' : f"Wrong Query Passed {key}"
            }, status=status.HTTP_206_PARTIAL_CONTENT)
    
    regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    email = query['email']

    if re.fullmatch(regex, email):
        username, domain = email.split('@')
        if domain != 'lnmiit.ac.in':
            return Response({
                'status' : "Error",
                'code' : 206,
                'message' : f"Enter a Valid LNMIIT Email Only"
            }, status = status.HTTP_206_PARTIAL_CONTENT)
        name = query['name']

        try:
            recommendObj = Recommend.objects.filter( email = email )
            if recommendObj.exists():
                books = []
                for obj in list(recommendObj.values()):
                    bookObj = Book.objects.get( id = obj['book_id'] )
                    books.append({
                        "id" : obj['id'],
                        'qty' : obj['quantity'],
                        "book_id" : bookObj.id,
                        "title" : bookObj.title,
                        "author" : bookObj.author,
                        "subject" : bookObj.subject,
                        "seller" : bookObj.seller.name,
                        "price" : bookObj.expected_price,
                        "medium" : bookObj.medium,
                        "price_denomination" : bookObj.price_denomination,
                    })
                
                return Response({
                        "status" : "Success",
                        "code":200,
                        "type" : "Recommended",
                        "typeCode" : 1,
                        "data":books,
                }, status=status.HTTP_200_OK)
            
            else:
                    return Response({
                        "status" : "error",
                        "code" : 204,
                        'Message' : "Something Went Wrong, Cant Fetch Recommended"
                    }, status=status.HTTP_204_NO_CONTENT)
        
        except Exception as e:
            print(e)
            return Response({
                'status' : 'Error', 
                'code' : 400,
                'message' : "Something Wrong with Recommend table"
            }, status = status.HTTP_400_BAD_REQUEST)

    else:
        return Response({
            'status' : "Error",
            'code' : 206,
            'message' : f"Enter a Valid Email"
        }, status=status.HTTP_206_PARTIAL_CONTENT)

@api_view(['POST'])
@permission_classes((AllowAny,))
@csrf_exempt
def getOrders(request):
    valid_keys = ['email', 'name']
    query = dict( request.data )
    query_parameters = list( query.keys() )

    for key in valid_keys:
        if key not in query_parameters:
            return Response({
                'status' : "Error",
                'code' : 400,
                'message' : f"Wrong Query Passed {key}"
            }, status=status.HTTP_206_PARTIAL_CONTENT)
    
    regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    email = query['email']

    if re.fullmatch(regex, email):
        username, domain = email.split('@')
        if domain != 'lnmiit.ac.in':
            return Response({
                'status' : "Error",
                'code' : 206,
                'message' : f"Enter a Valid LNMIIT Email Only"
            }, status = status.HTTP_206_PARTIAL_CONTENT)
        name = query['name']

        try:
            ordersObj = Order.objects.filter( email = email )
            if ordersObj.exists():
                books = []
                for obj in list(ordersObj.values()):
                    bookObj = Book.objects.get( id = obj['book_id'] )
                    books.append({
                        "id" : obj['id'],
                        'qty' : obj['quantity'],
                        "book_id" : bookObj.id,
                        "title" : bookObj.title,
                        "author" : bookObj.author,
                        "subject" : bookObj.subject,
                        "seller" : bookObj.seller.name,
                        "price" : bookObj.expected_price,
                        "medium" : bookObj.medium,
                        "price_denomination" : bookObj.price_denomination,
                    })
                
                return Response({
                        "status" : "Success",
                        "code":200,
                        "type" : "Purchase",
                        "typeCode" : 2,
                        "data":books,
                }, status=status.HTTP_200_OK)
            
            else:
                    return Response({
                        "status" : "error",
                        "code" : 204,
                        'Message' : "Something Went Wrong, Cant Fetch Purchases"
                    }, status=status.HTTP_204_NO_CONTENT)
        
        except Exception as e:
            print(e)
            return Response({
                'status' : 'Error', 
                'code' : 400,
                'message' : "Something Wrong with Purchases table"
            }, status = status.HTTP_400_BAD_REQUEST)

    else:
        return Response({
            'status' : "Error",
            'code' : 206,
            'message' : f"Enter a Valid Email"
        }, status=status.HTTP_206_PARTIAL_CONTENT)




def booksExcel(request):
    books = Book.objects.all()
    wb = Workbook()
    ws = wb.active

    ws["A1"] = "Book ID"
    ws["B1"] = "Title"
    ws["C1"] = "Subject"
    ws["D1"] = "Author"
    ws["E1"] = "Year of Publication"
    ws["F1"] = "Seller"
    ws["G1"] = "Publisher"
    ws["H1"] = "ISBN"
    ws["I1"] = "Medium"
    ws["J1"] = "Price_foreign_currency"
    ws["K1"] = "Price_indian_currency"
    ws["L1"] = "Price_denomination"
    ws["M1"] = "expected_price"
    ws["N1"] = "Book Link"

    row = 2

    for book in books:
        ws["A{}".format(row)] = book.id
        ws["B{}".format(row)] = book.title
        ws["C{}".format(row)] = book.subject
        ws["D{}".format(row)] = book.author
        ws["E{}".format(row)] = book.year_of_publication
        ws["F{}".format(row)] = book.seller.name
        ws["G{}".format(row)] = book.publisher
        ws["H{}".format(row)] = book.ISBN
        ws["I{}".format(row)] = book.medium
        ws["J{}".format(row)] = book.price_foreign_currency
        ws["K{}".format(row)] = book.price_indian_currency
        ws["L{}".format(row)] = book.price_denomination
        ws["M{}".format(row)] = book.expected_price
        ws["N{}".format(row)] = book.link
        row +=1
    
    response = HttpResponse(content=save_virtual_workbook(
        wb), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response['Content-Disposition'] = "attachment; filename=Master_List_All_Books.xlsx"

    print("File Created")
    return response

def booksRecommended(request):
    recommended = Recommend.objects.all()
    wb = Workbook()
    ws = wb.active

    ws["A1"] = "Book ID"
    ws["B1"] = "buyer"
    ws["C1"] = "email"
    ws["D1"] = "Title"
    ws["E1"] = "Subject"
    ws["F1"] = "Author"
    ws["G1"] = "Year of Publication"
    ws["H1"] = "Seller"
    ws["I1"] = "Publisher"
    ws["J1"] = "ISBN"
    ws["K1"] = "Medium"
    ws["L1"] = "Price_foreign_currency"
    ws["M1"] = "Price_indian_currency"
    ws["N1"] = "Price_denomination"
    ws["O1"] = "expected_price"
    ws["P1"] = "Book Link"

    row = 2

    for item in recommended:
        ws["A{}".format(row)] = item.id
        ws["B{}".format(row)] = item.buyer
        ws["C{}".format(row)] = item.email
        ws["D{}".format(row)] = item.title
        ws["E{}".format(row)] = item.subject
        ws["F{}".format(row)] = item.author
        ws["G{}".format(row)] = item.book.year_of_publication
        ws["H{}".format(row)] = item.seller_name
        ws["I{}".format(row)] = item.book.publisher
        ws["J{}".format(row)] = item.book.ISBN
        ws["K{}".format(row)] = item.book.medium
        ws["L{}".format(row)] = item.book.price_foreign_currency
        ws["M{}".format(row)] = item.book.price_indian_currency
        ws["N{}".format(row)] = item.book.price_denomination
        ws["O{}".format(row)] = item.book.expected_price
        ws["P{}".format(row)] = item.book.link
        row +=1
    
    response = HttpResponse(content=save_virtual_workbook(
        wb), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response['Content-Disposition'] = "attachment; filename=Recommended_List_All_Books.xlsx"

    print("File Created")
    return response


def booksSubjectRecommended(request, subject):
    return JsonResponse({
        "status" : "success",
        "code" : 200,
    })

from django.conf import settings
@api_view(['POST'])
@permission_classes([AllowAny,])
@csrf_exempt
def import_book_data(request):
    if len(request.FILES['files']) == 0:
        return JsonResponse({
            'status' : 'error',
            'code' : 301,
            'message' : " No File Found "
        })
    excel_obj = ExcelFileUpload.objects.create(excel_file_upload = request.FILES['files'])

    df = pd.read_csv(f"{settings.BASE_DIR}/media/{excel_obj.excel_file_upload}", encoding='UTF-8')
    books = df.values.tolist()
    for obj in books:
        title = obj[1]
        author = obj[2]
        subject = obj[3]
        year_of_publication = obj[4]
        edition = obj[5]
        ISBN = obj[6]
        print(ISBN)
        publisher = obj[7]
        seller_id = obj[8]
        if obj[9] is not None:
            price_foreign_currency = obj[9]
        else:
            price_foreign_currency = 0
        
        if obj[10] is not None:
            price_indian_currency = obj[10]
        else:
            price_indian_currency = 0
        link = obj[11]
        book = Book()
        seller = Seller.objects.get(id = seller_id)
        book.title = title
        book.subject = subject
        book.author = author
        book.seller = seller
        book.year_of_publication = year_of_publication
        book.edition = edition
        book.ISBN = ISBN
        book.publisher = publisher
        book.price_indian_currency = price_indian_currency
        book.price_foreign_currency = price_foreign_currency
        book.link = link
        if obj[12] == 'electronic':
            book.medium = obj[12]
        book.save()

        print(f'Book Added {book.id}')

    return JsonResponse({
        'status':'success',
        'code' : 200,
    })




# Admin Panel API
@api_view(['GET'])
@csrf_exempt
def getAdminCount(request, pk):
    recommendationObj = Recommend.objects.all()
    paperback = 0
    ebooks = 0
    cost = 0
    if pk > 2 and pk < 1:
        return JsonResponse({
            "status" : "error",
            "code" : 404,
            "Message" : "No Such Type Available"
        })
    for obj in recommendationObj:
        if pk == 1:
            if obj.book.medium == "PAPERBACK" and obj.recommended_to_library == True:
                paperback += 1
                if obj.book.price_denomination == 'INR' or obj.book.price_denomination == 'Rs':
                    cost += obj.book.expected_price
            elif obj.book.medium == "ELECTRONIC" and obj.recommended_to_library == True:
                if obj.book.price_denomination == 'INR' or obj.book.price_denomination == 'Rs':
                    cost += obj.book.expected_price
                ebooks += 1

        elif pk == 2:
            if obj.book.medium == "PAPERBACK" and obj.is_ordered == True:
                paperback += 1
                if obj.book.price_denomination == 'INR' or obj.book.price_denomination == 'Rs':
                    cost += obj.book.expected_price
            elif obj.book.medium == "ELECTRONIC" and obj.is_ordered == True:
                if obj.book.price_denomination == 'INR' or obj.book.price_denomination == 'Rs':
                    cost += obj.book.expected_price
                ebooks += 1
    return JsonResponse({
        "status" : "Success",
        "code" : 200,
        "data" : {
            "count" : paperback + ebooks,
            "paperbacks" : paperback,
            "ebooks" : ebooks,
            "cost" : cost
        }
    })

@api_view(['GET'])
@csrf_exempt
def booksActionsSeller(request, sellerid, type):
    if sellerid > 9 or sellerid < 1:
        return JsonResponse({
            "status" : "error",
            "code" : 404,
            "message" : "Seller ID Exceeded Permissible range"
        })
    if type > 2 or type < 1:
        return JsonResponse({
            "status" : "error",
            "code" : 404,
            "message" : "Type Exceeded Permissible range"
        })
    try:
        sellerObj = Seller.objects.get( id = sellerid )
    except:
        return JsonResponse({
            "status" : "error",
            "code" : 404,
            'Message' : "Seller Doesnt Exist"
        })
    if type == 1:
        object = Recommend.objects.filter( seller = sellerObj )
    else:
        object = Order.objects.filter( seller = sellerObj )
    count = [0, 0, 0, 0, 0, 0, 0]
    for obj in object:
        if obj.book.subject == 'CSE':
            count[0] += 1
        elif obj.book.subject == 'CCE':
            count[1] += 1
        elif obj.book.subject == 'ECE':
            count[2] += 1
        elif obj.book.subject == 'MME':
            count[3] += 1
        elif obj.book.subject == 'Physics':
            count[4] += 1
        elif obj.book.subject == 'HSS':
            count[5] += 1
        elif obj.book.subject == 'Mathematics':
            count[6] += 1
    print(count)
    return JsonResponse({
        "status" : "Success",
        "code" : 200,
        "data" : {
            "count" : count
        }
    })

@api_view(['GET'])
@csrf_exempt
def bookAction(request, type):
    if type > 2 or type < 1:
        return JsonResponse({
            "status" : "error",
            "code" : 404,
            "message" : "Type Exceeded Permissible range"
        })
    if type == 1:
        object = Recommend.objects.all()
    else:
        object = Order.objects.all()
    
    count = [0, 0, 0, 0, 0, 0, 0]
    for obj in object:
        if obj.book.subject == 'CSE':
            count[0] += 1
        elif obj.book.subject == 'CCE':
            count[1] += 1
        elif obj.book.subject == 'ECE':
            count[2] += 1
        elif obj.book.subject == 'MME':
            count[3] += 1
        elif obj.book.subject == 'Physics':
            count[4] += 1
        elif obj.book.subject == 'HSS':
            count[5] += 1
        elif obj.book.subject == 'Mathematics':
            count[6] += 1
    print(count)
    return JsonResponse({
        "status" : "Success",
        "code" : 200,
        "data" : {
            "count" : count
        }
    })

def recommendExcelApiAll(request, type):
    if type == 1:
        recommendObj = Recommend.objects.all()
    else :
        recommendObj = Order.objects.all()
    wb = Workbook()
    ws = wb.active

    ws["A1"] = "ID"
    ws["B1"] = "Name"
    ws["C1"] = "Email"
    ws["D1"] = "Date"
    ws["E1"] = "Title"
    ws["F1"] = "Subject"
    ws["G1"] = "Author"
    ws["H1"] = "Year of Publication"
    ws["I1"] = "Seller"
    ws["J1"] = "Publisher"
    ws["K1"] = "ISBN"
    ws["L1"] = "Medium"
    ws["M1"] = "Price_foreign_currency"
    ws["N1"] = "Price_indian_currency"
    ws["O1"] = "Denomination"
    ws["P1"] = "expected_price"
    ws["Q1"] = "Book Link"

    row = 2

    for obj in recommendObj:
        ws["A{}".format(row)] = obj.id
        ws["B{}".format(row)] = obj.buyer
        ws["C{}".format(row)] = obj.email
        ws["D{}".format(row)] = obj.created_at.date()
        ws["E{}".format(row)] = obj.book.title
        ws["F{}".format(row)] = obj.book.subject
        ws["G{}".format(row)] = obj.book.author
        ws["H{}".format(row)] = obj.book.year_of_publication
        ws["I{}".format(row)] = obj.book.seller.name
        ws["J{}".format(row)] = obj.book.publisher
        ws["K{}".format(row)] = obj.book.ISBN
        ws["L{}".format(row)] = obj.book.medium
        ws["M{}".format(row)] = obj.book.price_foreign_currency
        ws["N{}".format(row)] = obj.book.price_indian_currency
        ws["O{}".format(row)] = obj.book.price_denomination
        ws["P{}".format(row)] = obj.book.expected_price
        ws["Q{}".format(row)] = obj.book.link
        row += 1

    response = HttpResponse(content=save_virtual_workbook(
        wb), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response['Content-Disposition'] = "attachment; filename=Recommended_List_All_Books.xlsx"

    print("File Created")
    return response

def recommendExcelApiSeller(request, type, seller):
    sellerobj = Seller.objects.get(id = seller)
    if type == 1:
        recommendObj = Recommend.objects.filter( seller = sellerobj, recommended_to_library = True)
    else :
        recommendObj = Recommend.objects.filter( seller = sellerobj, is_ordered = True)
    wb = Workbook()
    ws = wb.active

    ws["A1"] = "ID"
    ws["B1"] = "Name"
    ws["C1"] = "Email"
    ws["D1"] = "Date"
    ws["E1"] = "Title"
    ws["F1"] = "Subject"
    ws["G1"] = "Author"
    ws["H1"] = "Year of Publication"
    ws["I1"] = "Seller"
    ws["J1"] = "Publisher"
    ws["K1"] = "ISBN"
    ws["L1"] = "Medium"
    ws["M1"] = "Price_foreign_currency"
    ws["N1"] = "Price_indian_currency"
    ws["O1"] = "Denomination"
    ws["P1"] = "expected_price"
    ws["Q1"] = "Book Link"

    row = 2

    for obj in recommendObj:
        ws["A{}".format(row)] = obj.id
        ws["B{}".format(row)] = obj.buyer
        ws["C{}".format(row)] = obj.email
        ws["D{}".format(row)] = obj.created_at.date()
        ws["E{}".format(row)] = obj.book.title
        ws["F{}".format(row)] = obj.book.subject
        ws["G{}".format(row)] = obj.book.author
        ws["H{}".format(row)] = obj.book.year_of_publication
        ws["I{}".format(row)] = obj.book.seller.name
        ws["J{}".format(row)] = obj.book.publisher
        ws["K{}".format(row)] = obj.book.ISBN
        ws["L{}".format(row)] = obj.book.medium
        ws["M{}".format(row)] = obj.book.price_foreign_currency
        ws["N{}".format(row)] = obj.book.price_indian_currency
        ws["O{}".format(row)] = obj.book.price_denomination
        ws["P{}".format(row)] = obj.book.expected_price
        ws["Q{}".format(row)] = obj.book.link

        row += 1

    response = HttpResponse(content=save_virtual_workbook(
        wb), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response['Content-Disposition'] = "attachment; filename=Recommended_List_Seller_Books.xlsx"

    print("File Created")
    return response

def recommendExcelApiSubject(request, type, subject):
    if type == 1:
        recommendObj = Recommend.objects.filter( recommended_to_library = True)
    else :
        recommendObj = Recommend.objects.filter( is_ordered = True)
    wb = Workbook()
    ws = wb.active

    ws["A1"] = "ID"
    ws["B1"] = "Name"
    ws["C1"] = "Email"
    ws["D1"] = "Date"
    ws["E1"] = "Title"
    ws["F1"] = "Subject"
    ws["G1"] = "Author"
    ws["H1"] = "Year of Publication"
    ws["I1"] = "Seller"
    ws["J1"] = "Publisher"
    ws["K1"] = "ISBN"
    ws["L1"] = "Medium"
    ws["M1"] = "Price_foreign_currency"
    ws["N1"] = "Price_indian_currency"
    ws["O1"] = "Denomination"
    ws["P1"] = "expected_price"
    ws["Q1"] = "Book Link"

    row = 2

    for obj in recommendObj:
        if obj.book.subject.lower() == subject.lower():
            ws["A{}".format(row)] = obj.id
            ws["B{}".format(row)] = obj.buyer
            ws["C{}".format(row)] = obj.email
            ws["D{}".format(row)] = obj.created_at.date()
            ws["E{}".format(row)] = obj.book.title
            ws["F{}".format(row)] = obj.book.subject
            ws["G{}".format(row)] = obj.book.author
            ws["H{}".format(row)] = obj.book.year_of_publication
            ws["I{}".format(row)] = obj.book.seller.name
            ws["J{}".format(row)] = obj.book.publisher
            ws["K{}".format(row)] = obj.book.ISBN
            ws["L{}".format(row)] = obj.book.medium
            ws["M{}".format(row)] = obj.book.price_foreign_currency
            ws["N{}".format(row)] = obj.book.price_indian_currency
            ws["O{}".format(row)] = obj.book.price_denomination
            ws["P{}".format(row)] = obj.book.expected_price
            ws["Q{}".format(row)] = obj.book.link
            row += 1

    response = HttpResponse(content=save_virtual_workbook(
        wb), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response['Content-Disposition'] = "attachment; filename=Recommended_List_Seller_And_Subject_Books.xlsx"

    print("File Created")
    return response

def recommendExcelApiSellerAndSubject(request, type, seller, subject):
    sellerobj = Seller.objects.get(id = seller)
    if type == 1:
        recommendObj = Recommend.objects.filter( seller = sellerobj, recommended_to_library = True)
    else :
        recommendObj = Recommend.objects.filter( seller = sellerobj, is_ordered = True)
    wb = Workbook()
    ws = wb.active

    ws["A1"] = "ID"
    ws["B1"] = "Name"
    ws["C1"] = "Email"
    ws["D1"] = "Date"
    ws["E1"] = "Title"
    ws["F1"] = "Subject"
    ws["G1"] = "Author"
    ws["H1"] = "Year of Publication"
    ws["I1"] = "Seller"
    ws["J1"] = "Publisher"
    ws["K1"] = "ISBN"
    ws["L1"] = "Medium"
    ws["M1"] = "Price_foreign_currency"
    ws["N1"] = "Price_indian_currency"
    ws["O1"] = "Denomination"
    ws["P1"] = "expected_price"
    ws["Q1"] = "Book Link"

    row = 2

    for obj in recommendObj:
        if obj.book.subject.lower() == subject.lower():
            ws["A{}".format(row)] = obj.id
            ws["B{}".format(row)] = obj.buyer
            ws["C{}".format(row)] = obj.email
            ws["D{}".format(row)] = obj.created_at.date()
            ws["E{}".format(row)] = obj.book.title
            ws["F{}".format(row)] = obj.book.subject
            ws["G{}".format(row)] = obj.book.author
            ws["H{}".format(row)] = obj.book.year_of_publication
            ws["I{}".format(row)] = obj.book.seller.name
            ws["J{}".format(row)] = obj.book.publisher
            ws["K{}".format(row)] = obj.book.ISBN
            ws["L{}".format(row)] = obj.book.medium
            ws["M{}".format(row)] = obj.book.price_foreign_currency
            ws["N{}".format(row)] = obj.book.price_indian_currency
            ws["O{}".format(row)] = obj.book.price_denomination
            ws["P{}".format(row)] = obj.book.expected_price
            ws["Q{}".format(row)] = obj.book.link
            row += 1

    response = HttpResponse(content=save_virtual_workbook(
        wb), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response['Content-Disposition'] = "attachment; filename=Recommended_List_Seller_And_Subject_Books.xlsx"

    print("File Created")
    return response


