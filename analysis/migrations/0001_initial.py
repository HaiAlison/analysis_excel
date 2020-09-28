# Generated by Django 3.0.7 on 2020-09-16 11:17

from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion
import django.utils.timezone


class Migration(migrations.Migration):

    initial = True

    operations = [
        migrations.CreateModel(
            name='Import',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('link_to_specs', models.CharField(max_length=100, unique=True)),
                ('date_imported', models.DateTimeField(default=django.utils.timezone.now)),
            ],
        ),
    ]