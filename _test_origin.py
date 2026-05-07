"""Test: set page.active=1 then plotxy, and try win32com raw COM."""
import originpro as op, numpy as np
from originpro.config import po

op.set_show(True)
op.new()

x  = np.linspace(100.0, 3500.0, 10)
yu = 1.2 - 0.0001 * x

book = op.new_book('w', lname='TestBook')
wks  = book[0]
wks.name = 'Sheet1'
wks.cols = 2
wks.from_list(0, x.tolist(), lname='I', axis='X')
wks.from_list(1, yu.tolist(), lname='V', axis='Y')
bname = book.name

graph = op.new_graph(lname='TG')
gl = graph[0]

# Activate the layer
gl.activate()

# Explicitly set page.active = 1
op.lt_exec("page.active = 1;")
active_layer = op.lt_int("page.active")
active_page  = op.get_lt_str("page.name$")
print(f"Active page: {active_page}, active layer: {active_layer}")

# Try plotxy with explicit x,y col references
op.lt_exec(f"plotxy ([{bname}]Sheet1!col(A), [{bname}]Sheet1!col(B)) type:=200;")
n = gl.get_int("numplots")
print("numplots after plotxy:", n)

# Also inspect the po object
print("\npo type:", type(po).__name__)
print("po attrs sample:", [m for m in dir(po) if not m.startswith('_')][:20])

# Try raw Execute on the page COM object
page_obj = graph.obj
print("\npage_obj type:", type(page_obj).__name__)
print("page_obj attrs:", [m for m in dir(page_obj) if not m.startswith('_')][:30])
