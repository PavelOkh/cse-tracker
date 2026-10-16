"""DPD parcel tracing. No Streamlit dependency; no consuming/confirm calls."""
from datetime import datetime, timezone
import time

LABELS = {'NewOrderByClient':'Заказ создан клиентом','NewOrderByDPD':'Заказ создан DPD',
 'NotDone':'Заказ отменён','OnTerminalPickup':'На терминале отправления',
 'OnRoad':'В пути','OnTerminal':'На транзитном терминале',
 'OnTerminalDelivery':'На терминале доставки','Delivering':'Передано на доставку',
 'Delivered':'Доставлено','Lost':'Утеряно','Problem':'Проблемная ситуация',
 'ReturnedFromDelivery':'Возвращено с доставки'}


def timestamp(value):
    if not value:
        raise ValueError('В ответе DPD отсутствует transitionTime')
    dt = value if isinstance(value, datetime) else datetime.fromisoformat(str(value).replace('Z','+00:00'))
    # DPD examples use local timestamps without an offset. Do not invent an offset in output.
    return dt.replace(tzinfo=timezone.utc).timestamp() if dt.tzinfo is None else dt.timestamp()


def normalize_states(number, response):
    """One row per parcel and return direction. Never collapse partial deliveries."""
    states = response.get('states') or []
    if isinstance(states, dict):
        states = [states]
    groups = {}
    for state in states:
        key = (str(state.get('dpdOrderNr') or number), str(state.get('dpdParcelNr') or state.get('clientParcelNr') or ''),
               str(state.get('pickupDate') or ''), str(state.get('isReturn') or False).lower())
        groups.setdefault(key, []).append(state)
    rows = []
    for (order, parcel, pickup, returning), events in groups.items():
        if not parcel:
            raise ValueError('DPD не вернул идентификатор места; проверьте исходный ответ')
        current = max(events, key=lambda s: timestamp(s.get('transitionTime')))
        delivered = [s for s in events if s.get('newState') == 'Delivered']
        delivery = max(delivered, key=lambda s: timestamp(s.get('transitionTime'))) if delivered else {}
        code = current.get('newState','')
        rows.append({'Накладная':number,'Заказ DPD':order,'Место DPD':parcel,'Дата приёма':pickup,
            'Возврат':returning == 'true','Код статуса':code,'Статус':LABELS.get(code,code or 'Не указан'),
            'Дата посл. статуса':str(current.get('transitionTime') or ''),
            'Дата доставки':str(delivery.get('transitionTime') or ''),
            'План доставки':str(current.get('planDeliveryDate') or ''),
            'Город':current.get('terminalCity') or '', 'Инцидент':current.get('incidentName') or '',
            'Связанный заказ':current.get('dpdOrderReNr') or '', 'Ошибка':''})
    return rows or [{'Накладная':number,'Статус':'История не найдена','Ошибка':''}]


class DPDTracker:
    def __init__(self, client_number, client_key, test=False):
        import requests
        from zeep import Client, Settings
        from zeep.transports import Transport
        if not str(client_number).isdigit() or not client_key:
            raise ValueError('Укажите номер клиента DPD (цифры) и API-ключ')
        self.auth = {'clientNumber':int(client_number), 'clientKey':client_key}
        host = 'wstest.dpd.ru' if test else 'ws.dpd.ru'
        session = requests.Session()
        # Only retry read-only tracing calls on transient HTTP failures.
        from requests.adapters import HTTPAdapter
        from urllib3.util.retry import Retry
        session.mount('https://', HTTPAdapter(max_retries=Retry(total=2, backoff_factor=1,
            status_forcelist=[429,502,503,504], allowed_methods=['GET','POST'], respect_retry_after_header=False)))
        self.client = Client(f'https://{host}/services/tracing?wsdl',
            transport=Transport(session=session, timeout=20, operation_timeout=30),
            settings=Settings(strict=True, force_https=True))
        # WSDLs may advertise legacy HTTP endpoints. Credentials must travel over TLS.
        self.service = self.client.create_service(self.client.service._binding.name,
                                                   f'https://{host}/services/tracing')

    def check(self, number, pickup_year=None):
        from zeep.helpers import serialize_object
        from zeep.exceptions import Fault
        request = {'auth':self.auth,'dpdOrderNr':str(number)}
        if pickup_year is not None:
            request['pickupYear'] = int(pickup_year)
        try:
            response = self.service.getStatesByDPDOrder(request=request)
            return normalize_states(str(number), serialize_object(response) or {})
        except Fault as exc:
            # Do not expose raw SOAP requests or server fault bodies containing credentials.
            msg = str(exc).lower()
            code = next((c for c in ['no-data-found','no-setting','auth-error'] if c in msg), 'SOAP Fault')
            explanations = {'no-data-found':'История не найдена: проверьте номер и год, возможна ещё не выполненная приёмка',
                'no-setting':'У клиента DPD не включена услуга отслеживания', 'auth-error':'Ошибка авторизации DPD'}
            return [{'Накладная':str(number),'Статус':explanations.get(code,'Ошибка SOAP DPD'),'Ошибка':code}]
        except Exception as exc:
            return [{'Накладная':str(number),'Статус':'Ошибка проверки','Ошибка':type(exc).__name__}]

    def check_many(self, numbers, pickup_year=None, progress=None):
        rows = []
        numbers = list(dict.fromkeys(str(n).strip() for n in numbers if str(n).strip()))
        for i, number in enumerate(numbers):
            rows.extend(self.check(number, pickup_year))
            if progress:
                progress((i+1)/len(numbers))
            if i+1 < len(numbers):
                time.sleep(0.3)
        return rows
