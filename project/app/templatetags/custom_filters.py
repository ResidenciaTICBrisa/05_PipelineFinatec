from django import template

register = template.Library()

@register.filter(name='first_letters')
def first_letters(email):
    parts = email.split('@')
    username = parts[0]
    domain = parts[1]
    asteriscos = ''
    for i in range(len(username)-5):
        asteriscos+='*'

    if len(username) <= 2:
        return username
    else:
        return username[:2] + asteriscos + username[-2:] + '@' + domain
