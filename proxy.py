from ConfigParser import ConfigParser, NoSectionError, NoOptionError
from collections import namedtuple
from hid import device
from locale import setlocale, LC_ALL
from logging import getLogger, INFO, StreamHandler, Formatter, warning, info, error
from os.path import splitext, abspath, exists
from serial import Serial
from struct import unpack
from sys import stdout
from time import sleep, time


SERIAL_SECTION = 'serial'
PORT_NUMBER = 'port_number'
INVALID_PORT_FORMAT = 'Please specify port number to be used inside %s'

Config = namedtuple('Config', (
    'port_name',
))

QUERY_INTERVAL_IN_SECONDS = 1

VENDOR_ID = 0
PRODUCT_ID = 1

REQUEST = [
    0x00, 0x0F, 0x03, 0x55, 0x00, 0x55, 0xFF, 0xFF, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
]
RESPONSE_BUFFER_SIZE = 64
RESPONSE_FORMAT = '>Lbhhhhbbhhhhhhh12s23s'

ResponseFormat = namedtuple('ResponseFormat', (
    'prefix',
    'state',
    'charge',
    'timer',
    'milli_voltage',
    'milli_current',
    'external_temperature',
    'internal_temperature',
    'error_notice',
    'milli_voltage_cell_1',
    'milli_voltage_cell_2',
    'milli_voltage_cell_3',
    'milli_voltage_cell_4',
    'milli_voltage_cell_5',
    'milli_voltage_cell_6',
    'unknown_values',
    'zeros',
))

RESPONSE_PREFIX = 0x0f225500

STATE_MESSAGES = {
    1: 'running',
    2: 'user stop',
    3: 'imax finish',
}

OpenFormat = namedtuple('OpenFormat', (
    'channel_number',
    'state',
    'timer',
    'voltage',
    'current',
    'charge',
    'power',
    'external_temperature',
    'internal_temperature',
    'voltage_cell_1',
    'voltage_cell_2',
    'voltage_cell_3',
    'voltage_cell_4',
    'voltage_cell_5',
    'voltage_cell_6',
    'check_sum',
))

OPEN_FORMAT_CHANNEL_NUMBER = 1
OPEN_FORMAT_STATE = 1
OPEN_FORMAT_CHECK_SUM = 0


previous_state = None


def parse(data):
    global previous_state

    packet = ''.join(chr(value) for value in data)
    structure = unpack(RESPONSE_FORMAT, packet)
    response_format = ResponseFormat(*structure)

    if response_format.prefix != RESPONSE_PREFIX:
        warning('Unknown prefix: %08X' % response_format.prefix)
    if response_format.error_notice:
        warning('Error notice: %08X' % response_format.error_notice)
    if any(char != '\0' for char in response_format.zeros):
        hexes = ''.join('%02X' % ord(char) for char in response_format.zeros)
        warning('Non zero value in the tail: %s' % hexes)

    if previous_state != response_format.state:
        previous_state = response_format.state
        info('New state: %s' % STATE_MESSAGES.get(response_format.state, 'unknown'))

    return response_format


def create_open_format(response_format):
    return OpenFormat(
        channel_number=OPEN_FORMAT_CHANNEL_NUMBER,
        state=OPEN_FORMAT_STATE,
        timer=response_format.timer,
        voltage=response_format.milli_voltage / 1000.0,
        current=response_format.milli_current / 1000.0,
        charge=response_format.charge,
        power=response_format.milli_voltage / 1000.0 * response_format.milli_current / 1000.0,
        external_temperature=response_format.external_temperature,
        internal_temperature=response_format.internal_temperature,
        voltage_cell_1=response_format.milli_voltage_cell_1 / 1000.0,
        voltage_cell_2=response_format.milli_voltage_cell_2 / 1000.0,
        voltage_cell_3=response_format.milli_voltage_cell_3 / 1000.0,
        voltage_cell_4=response_format.milli_voltage_cell_4 / 1000.0,
        voltage_cell_5=response_format.milli_voltage_cell_5 / 1000.0,
        voltage_cell_6=response_format.milli_voltage_cell_6 / 1000.0,
        check_sum=OPEN_FORMAT_CHECK_SUM,
    )


def serialize(open_format):
    payload = ';'.join(
        ('%g' % value).replace('.', ',')
        for value in open_format
    )
    return '$%s\r\n' % payload


def resend(incoming_values, serial_port):
    response_format = parse(incoming_values)
    open_format = create_open_format(response_format)
    output_data = serialize(open_format)
    serial_port.write(output_data)


def get_hid_device():
    hid_device = device()
    hid_device.open(VENDOR_ID, PRODUCT_ID)
    return hid_device


def get_config_path():
    name, _ext = splitext(abspath(__file__))
    return '%s.ini' % name


def create_dummy_config(config_path):
    config = ConfigParser()
    config.add_section(SERIAL_SECTION)
    config.set(SERIAL_SECTION, PORT_NUMBER, '')
    with open(config_path, 'w') as output:
        config.write(output)


def read_config(config_path):
    config = ConfigParser()
    config.read(config_path)
    try:
        port_string = config.get(SERIAL_SECTION, PORT_NUMBER)
    except (NoSectionError, NoOptionError):
        error('Invalid config file %s' % config_path)
        info('You can delete it to create new one')
        raise
    try:
        port_number = int(port_string)
    except ValueError:
        error(INVALID_PORT_FORMAT % config_path)
        raise
    return Config(
        port_name='COM%d' % port_number,
    )


def get_config():
    path = get_config_path()
    if not exists(path):
        create_dummy_config(path)
    return read_config(path)


def setup_logger():
    setlocale(LC_ALL, '')
    formatter = Formatter(
        fmt='[%(asctime)s] %(levelname)8s: %(message)s',
        datefmt='%x %X')

    stream_handler = StreamHandler(stdout)
    stream_handler.setFormatter(fmt=formatter)

    logger = getLogger()
    logger.setLevel(INFO)
    logger.addHandler(stream_handler)


def run():
    setup_logger()

    config = get_config()
    hid_device = None
    serial_port = Serial(config.port_name)

    next_start = time()
    while True:
        next_start += QUERY_INTERVAL_IN_SECONDS

        try:
            if hid_device is None:
                hid_device = get_hid_device()
            hid_device.write(REQUEST)
            incoming_values = hid_device.read(RESPONSE_BUFFER_SIZE, QUERY_INTERVAL_IN_SECONDS * 1000)
        except IOError:
            error('can\'t read from USB HID')
            if hid_device is not None:
                try:
                    hid_device.close()
                except IOError:
                    pass
            hid_device = None
        else:
            if len(incoming_values) < RESPONSE_BUFFER_SIZE:
                error(
                    '%d values have been received instead of %d' % (
                        len(incoming_values),
                        RESPONSE_BUFFER_SIZE,
                    )
                )
                continue
            resend(incoming_values, serial_port)

        now = time()
        if now > next_start:
            next_start = now
        else:
            sleep(next_start - now)


if __name__ == '__main__':
    run()
