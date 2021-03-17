import os
import unittest
import subprocess
import time
import signal
from oodocument import oodocument

OOPORT = 8001
OOHOST ="0.0.0.0"
OOCMD = 'libreoffice --headless --nofirststartwizard --accept="socket,host={},port={};urp"'.format(OOHOST, OOPORT)

dir_path = os.path.dirname(os.path.realpath(__file__))

def start_office_instance(cmd):
    """
    Starts Libre/Open Office with a listening socket.
    @type  office: string
    @param office: Libre/Open Office startup string
    """
    # Start OpenOffice.org and report any errors that occur.
    try:
        p = subprocess.Popen(cmd, stdout=subprocess.PIPE, shell=True, preexec_fn=os.setsid)
    except OSError as ose:
        raise OSError(ose)
    return p


class OODocumentTest(unittest.TestCase):
    def setUp(self):
        self.proc = start_office_instance(OOCMD)

    def test_set_document(self):
        time.sleep(1)
        filename = os.path.join(dir_path, 'test.docx')
        oo = oodocument(filename, host=OOHOST, port=OOPORT)
        self.assertEqual(str(oo), filename)
        self.assertIn('pyuno', str(oo.document))
        oo.dispose()

    def tearDown(self):
        os.killpg(os.getpgid(self.proc.pid), signal.SIGTERM)
        self.proc.wait()

if __name__ == "__main__":
    unittest.main()
