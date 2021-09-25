# Generated by Django 3.2.7 on 2021-09-22 16:05

import django.core.validators
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='Book',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('title', models.CharField(blank=True, max_length=600, null=True)),
                ('author', models.CharField(blank=True, max_length=800, null=True)),
                ('subject', models.CharField(blank=True, max_length=200, null=True)),
                ('year_of_publication', models.PositiveIntegerField(blank=True, null=True)),
                ('edition', models.CharField(blank=True, max_length=50, null=True)),
                ('ISBN', models.CharField(blank=True, max_length=200, null=True)),
                ('publisher', models.CharField(blank=True, max_length=500, null=True)),
                ('medium', models.CharField(choices=[('PAPERBACK', 'paperback'), ('ELECTRONIC', 'electronic')], default='PAPERBACK', max_length=200)),
                ('price_foreign_currency', models.CharField(blank=True, max_length=200, null=True)),
                ('price_indian_currency', models.CharField(blank=True, max_length=200, null=True)),
                ('price', models.FloatField(blank=True, null=True)),
                ('price_denomination', models.CharField(blank=True, default='INR', max_length=20, null=True)),
                ('discount', models.PositiveSmallIntegerField(default=25, validators=[django.core.validators.MaxValueValidator(100), django.core.validators.MinValueValidator(0)])),
                ('expected_price', models.FloatField(blank=True, null=True, validators=[django.core.validators.MinValueValidator(0)])),
                ('link', models.URLField(blank=True, max_length=500, null=True)),
                ('visible', models.BooleanField(default=True)),
                ('expired', models.BooleanField(default=False)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('updated_at', models.DateTimeField(auto_now=True)),
                ('image', models.URLField(blank=True, null=True)),
            ],
            options={
                'ordering': ['id'],
            },
        ),
        migrations.CreateModel(
            name='Seller',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(blank=True, max_length=200, null=True)),
                ('contact_no', models.CharField(blank=True, max_length=10, null=True, unique=True)),
                ('email', models.EmailField(blank=True, max_length=200, null=True)),
                ('logo', models.ImageField(blank=True, null=True, upload_to='book_sellers_logo')),
            ],
        ),
        migrations.CreateModel(
            name='UserType',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(blank=True, max_length=200, null=True)),
                ('code', models.IntegerField(blank=True, default=4, null=True)),
            ],
        ),
        migrations.CreateModel(
            name='Recommend',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(blank=True, max_length=200, null=True)),
                ('email', models.EmailField(max_length=200, null=True)),
                ('recommended_to_library', models.BooleanField(default=False)),
                ('quantity', models.IntegerField(blank=True, default=1, null=True)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('updated_at', models.DateTimeField(auto_now=True)),
                ('book', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='recommend', to='api.book')),
                ('seller', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='recommend', to='api.seller')),
                ('usertype', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='recommend', to='api.usertype')),
            ],
        ),
        migrations.CreateModel(
            name='Order',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(blank=True, max_length=200, null=True)),
                ('email', models.EmailField(max_length=200, null=True)),
                ('is_ordered', models.BooleanField(default=False)),
                ('quantity', models.IntegerField(blank=True, default=1, null=True)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('updated_at', models.DateTimeField(auto_now=True)),
                ('book', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='order', to='api.book')),
                ('seller', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='order', to='api.seller')),
                ('usertype', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='order', to='api.usertype')),
            ],
        ),
        migrations.AddField(
            model_name='book',
            name='seller',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, to='api.seller'),
        ),
    ]
