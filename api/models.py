import re
from django.db import models
from django.db import models
from django.core.validators import MaxValueValidator, MinValueValidator
from django.db.models import base

# Create your models here.
class Seller(models.Model):
    """
    The Details of a Book Seller
    """
    name = models.CharField(max_length=200, blank=True, null = True)
    contact_no = models.CharField(max_length = 10, blank=True, null=True)
    email=models.EmailField(max_length=200,null=True, blank=True)
    logo = models.ImageField(
        upload_to='book_sellers_logo', null=True, blank=True)
    def __str__(self):
        return f"{self.name}-{self.id}"


class Book(models.Model):
    """
    Model to save a Publisher 's Book
    Can be changed or updated
    """
    MEDIUM_CHOICES = (
        ("PAPERBACK", "paperback"),
        ("ELECTRONIC", 'electronic'), 
    )
    title = models.CharField(max_length=600, blank=True, null=True)
    author = models.CharField(max_length=800, null=True, blank=True)
    subject = models.CharField(max_length=200, blank=True, null=True)
    year_of_publication = models.PositiveIntegerField(blank=True, null=True)
    edition = models.CharField(max_length=50, null=True, blank=True)
    ISBN = models.CharField(max_length=200, null=True, blank=True)
    publisher = models.CharField(max_length=500, null = True, blank=True)

    seller = models.ForeignKey(Seller, on_delete=models.CASCADE, null = True, blank=True)

    medium = models.CharField(max_length=200, choices=MEDIUM_CHOICES, default= "PAPERBACK")
    price_foreign_currency = models.CharField(max_length=200, blank=True, null=True)
    price_indian_currency = models.CharField(max_length=200,  blank=True, null=True)
    price = models.FloatField(blank=True, null=True)
    price_denomination = models.CharField(max_length=20, blank=True, null=True, default='INR')
    discount = models.PositiveSmallIntegerField(
        default=25, validators=[MaxValueValidator(100), MinValueValidator(0)])
    expected_price = models.FloatField(
        validators=[MinValueValidator(0)], blank=True, null=True)
    
    link = models.URLField(max_length=500, blank=True, null=True)
    # For banning/unbanning products.
    visible = models.BooleanField(default=True)
    # For expiring products after given expiry period.
    expired = models.BooleanField(default=False)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    image = models.URLField(blank=True, null=True)

    class Meta:
        ordering = ['id']

    def __str__(self):
        return f"{self.title} + {self.seller.name}"
    
    def save(self, *args, **kwargs):
        self.price_indian_currency = str(self.price_indian_currency).replace(',', '')
        output = re.findall(r'\d+', self.price_indian_currency)
        if len(output) == 0:  
            self.price_foreign_currency = str(self.price_foreign_currency).replace(',', '')
            output = re.findall(r'\d+', self.price_foreign_currency)
            if len( output ) == 0:
                value = 0
                demo = 'INR'
            else:
                value = output[0]
                result = re.findall(r'[a-zA-Z]+', self.price_foreign_currency)
                if len(result) > 0:
                    demo = result[0]
                else:
                    demo = 'FC'
        else:
            value = output[0]
            demo = 'INR'
        self.price = float(value)
        if demo:
            self.price_denomination = demo


        if self.price is not None:
            self.expected_price = round(float(self.price) * (1-((float(self.discount))/100)))

        if self.pk is None and str(self.ISBN).isdecimal():
            self.image = "http://covers.openlibrary.org/b/isbn/"+str(int(self.ISBN))+"-L.jpg"
            self.thumbnail = "http://covers.openlibrary.org/b/isbn/"+str(int(self.ISBN))+"-S.jpg"
        super().save(*args, **kwargs)

class UserType(models.Model):
    name = models.CharField(max_length=200, blank=True, null=True)
    code = models.IntegerField(default=4, blank=True, null=True )

    def __str__(self):
        return f"{self.name}-{self.code}"

class Recommend(models.Model):
    """
    Model to save a user's recommendation. 
    Users can add/delete recommendation from their cart.
    """
    book = models.ForeignKey(
        Book,
        related_name="recommend",
        on_delete=models.SET_NULL,
        null=True,
        blank=True
    )
    seller = models.ForeignKey(
        Seller,
        related_name="recommend",
        on_delete=models.CASCADE,
        null=True,
        blank=True
    )
    usertype = models.ForeignKey(UserType, related_name='recommend', on_delete=models.CASCADE, null=True, blank=True)

    name=models.CharField(max_length=200, blank=True, null = True)
    email=models.EmailField(max_length=200,null=True)
    recommended_to_library = models.BooleanField(default=False)
    quantity=models.IntegerField(default=1, null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.buyer} - {self.email}"
    
    def save(self, *args, **kwargs):
        if self.seller is None:
            self.seller = self.book.seller
        super().save(*args, **kwargs)



class Order(models.Model):
    """
    Model to save a user's Orders for personal purchase. 
    Users can add/delete recommendation from their cart.
    """
    book = models.ForeignKey(
        Book,
        related_name="order",
        on_delete=models.SET_NULL,
        null=True,
        blank=True
    )
    seller = models.ForeignKey(
        Seller,
        related_name="order",
        on_delete=models.CASCADE,
        null=True,
        blank=True
    )
    usertype = models.ForeignKey(UserType, related_name='order', on_delete=models.CASCADE, null=True, blank=True)

    name=models.CharField(max_length=200, blank=True, null = True)
    email=models.EmailField(max_length=200,null=True)
    is_ordered = models.BooleanField(default=False)
    quantity=models.IntegerField(default=1, null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def save(self, *args, **kwargs):
        if self.seller is None:
            self.seller = self.book.seller
        super().save(*args, **kwargs)

    def __str__(self):
        return f"{self.name} - {self.email}"


class ExcelFileUpload(models.Model):
    excel_file_upload = models.FileField(upload_to='excel')