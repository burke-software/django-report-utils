""" Functioned to introspect a model """
from django.contrib.contenttypes.models import ContentType
from django.db.models.fields import FieldDoesNotExist
from django.conf import settings
import inspect

def isprop(v):
    return isinstance(v, property)

def get_properties_from_model(model_class):
    """ Show properties from a model """
    properties = []
    attr_names = [name for (name, value) in inspect.getmembers(model_class, isprop)]
    for attr_name in attr_names:
        if attr_name.endswith('pk'):
            attr_names.remove(attr_name)
        else:
            properties.append(dict(label=attr_name, name=attr_name.strip('_').replace('_',' ')))
    return sorted(properties, key=lambda k: k['label'])


def get_relation_fields_from_model(model_class):
    """ Get related fields (m2m, FK, and reverse FK) """
    relation_fields = []
    all_fields_names = model_class._meta.get_all_field_names()
    for field_name in all_fields_names:
        field = model_class._meta.get_field_by_name(field_name)
        # get_all_field_names will return the same field
        # both with and without _id. Ignore the duplicate.
        if field_name[-3:] == '_id' and field_name[:-3] in all_fields_names:
            continue
        if field[3] or not field[2] or hasattr(field[0], 'related'):
            field[0].field_name = field_name
            relation_fields += [field[0]]
    return relation_fields


def get_direct_fields_from_model(model_class):
    """ Direct, not m2m, not FK """
    direct_fields = []
    all_fields_names = model_class._meta.get_all_field_names()
    for field_name in all_fields_names:
        field = model_class._meta.get_field_by_name(field_name)
        if field[2] and not field[3] and not hasattr(field[0], 'related'):
            direct_fields += [field[0]]
    return direct_fields


def get_custom_fields_from_model(model_class):
    """ django-custom-fields support """
    if 'custom_field' in settings.INSTALLED_APPS:
        from custom_field.models import CustomField
        try:
            content_type = ContentType.objects.get(
                model=model_class._meta.model_name,
                app_label=model_class._meta.app_label)
        except ContentType.DoesNotExist:
            content_type = None
        custom_fields = CustomField.objects.filter(content_type=content_type)
        return custom_fields


def get_model_from_path_string(root_model, path):
    """ Return a model class for a related model
    root_model is the class of the initial model
    path is like foo__bar where bar is related to foo
    """
    for path_section in path.split('__'):
        if path_section:
            try:
                field = root_model._meta.get_field_by_name(path_section)
            except FieldDoesNotExist:
                return root_model
            if field[2]:
                if hasattr(field[0], 'related'):
                    try:
                        root_model = field[0].related.parent_model()
                    except AttributeError:
                        root_model = field[0].related.model
            else:
                if hasattr(field[0], 'related_model'):
                    root_model = field[0].related_model
                else:
                    root_model = field[0].model
    return root_model
