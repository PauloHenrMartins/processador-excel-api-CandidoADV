import json
import datetime

class DateTimeEncoder(json.JSONEncoder):
    """Codificador JSON personalizado para lidar com objetos datetime."""
    def default(self, obj):
        if isinstance(obj, datetime.datetime):
            return obj.isoformat()
        return json.JSONEncoder.default(self, obj)
