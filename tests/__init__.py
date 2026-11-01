import logging

# Ожидаемые в тестах предупреждения (ретраи OSRM, битый state) не засоряют вывод
logging.disable(logging.CRITICAL)
