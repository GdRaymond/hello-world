import os
from flask import Flask
from flask_script import Manager, Shell
from flask_sqlalchemy import SQLAlchemy

basedir=os.path.abspath(os.path.dirname(__file__))
app=Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI']='sqlite:///'+os.path.join(basedir,'data.sqlite')
app.config['SQLALCHEMY_COMMIT_ON_TEARDOWN']=True
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
manager=Manager(app)

db=SQLAlchemy(app)

class Order_info(db.Model):
    __tablename__='order_info'
    id=db.Column(db.Integer,primary_key=True)
    tis_no=db.Column(db.String(13))
    abm_no=db.Column(db.String(11))
    fty=db.Column(db.String(20))
    style_no=db.Column(db.String(20))
    commodity=db.Column(db.String(50))
    colour=db.Column(db.String(20))
    freight_way=db.Column(db.String(10))
    etd=db.Column(db.Date)
    eta=db.Column(db.Date)
    clear=db.Column(db.Date)
    in_store=db.Column(db.Date)
    order_date=db.Column(db.Date)
    #status: new,shipped,finished
    status=db.Column(db.String(10))
    #source: packing list-sheet name,'order form'
    source=db.Column(db.String(100),default='order form')
    price=db.Column(db.Float,default=0.0)
    size_1=db.Column(db.Integer,default=0)
    size_2=db.Column(db.Integer,default=0)
    size_3=db.Column(db.Integer,default=0)
    size_4=db.Column(db.Integer,default=0)
    size_5=db.Column(db.Integer,default=0)
    size_6=db.Column(db.Integer,default=0)
    size_7=db.Column(db.Integer,default=0)
    size_8=db.Column(db.Integer,default=0)
    size_9=db.Column(db.Integer,default=0)
    size_10=db.Column(db.Integer,default=0)
    size_11=db.Column(db.Integer,default=0)
    size_12=db.Column(db.Integer,default=0)
    size_13=db.Column(db.Integer,default=0)
    size_14=db.Column(db.Integer,default=0)
    size_15=db.Column(db.Integer,default=0)
    size_16=db.Column(db.Integer,default=0)
    size_17=db.Column(db.Integer,default=0)
    size_18=db.Column(db.Integer,default=0)
    size_19=db.Column(db.Integer,default=0)
    size_20=db.Column(db.Integer,default=0)
    size_21=db.Column(db.Integer,default=0)
    size_22=db.Column(db.Integer,default=0)
    size_23=db.Column(db.Integer,default=0)
    size_24=db.Column(db.Integer,default=0)
    size_25=db.Column(db.Integer,default=0)
    size_26=db.Column(db.Integer,default=0)
    size_27=db.Column(db.Integer,default=0)
    size_28=db.Column(db.Integer,default=0)
    size_29=db.Column(db.Integer,default=0)
    size_30=db.Column(db.Integer,default=0)

    def __repr__(self):
        return('<Order:%s/%s/%s>'%(self.tis_no,self.style_no,self.colour))

class Packings(db.Model):
    __tablename__='packings'
    id=db.Column(db.Integer,primary_key=True)
    tis_no=db.Column(db.String(13))
    abm_no=db.Column(db.String(11))
    fty=db.Column(db.String(20))
    style_no=db.Column(db.String(20))
    commodity=db.Column(db.String(50))
    price=db.Column(db.Float,default=0.0)
    invoice_no=db.Column(db.String(50))
    total_qty=db.Column(db.Integer,default=0)
    total_carton=db.Column(db.Integer,default=0)
    total_gw=db.Column(db.Float,default=0.0)
    total_volume=db.Column(db.Float,default=0.0)
    invoice_date=db.Column(db.Date)
    #source: packing list-sheet name,'order form'
    source=db.Column(db.String(100))
    receives=db.relationship('Actual_qty',backref='packing',lazy='dynamic')
    details=db.relationship('Detail_carton',backref='packing',lazy='dynamic')

    def __repr__(self):
        return('<Packing list :%s/%s/%s>'%(self.tis_no,self.style_no,source))

class Actual_qty(db.Model):
    __tablename__='actual_qty'
    id=db.Column(db.Integer,primary_key=True)
    packing_id=db.Column(db.Integer,db.ForeignKey('packings.id'))
    colour=db.Column(db.String(20))
    total_qty=db.Column(db.Integer,default=0)
    size_1=db.Column(db.Integer,default=0)
    size_2=db.Column(db.Integer,default=0)
    size_3=db.Column(db.Integer,default=0)
    size_4=db.Column(db.Integer,default=0)
    size_5=db.Column(db.Integer,default=0)
    size_6=db.Column(db.Integer,default=0)
    size_7=db.Column(db.Integer,default=0)
    size_8=db.Column(db.Integer,default=0)
    size_9=db.Column(db.Integer,default=0)
    size_10=db.Column(db.Integer,default=0)
    size_11=db.Column(db.Integer,default=0)
    size_12=db.Column(db.Integer,default=0)
    size_13=db.Column(db.Integer,default=0)
    size_14=db.Column(db.Integer,default=0)
    size_15=db.Column(db.Integer,default=0)
    size_16=db.Column(db.Integer,default=0)
    size_17=db.Column(db.Integer,default=0)
    size_18=db.Column(db.Integer,default=0)
    size_19=db.Column(db.Integer,default=0)
    size_20=db.Column(db.Integer,default=0)
    size_21=db.Column(db.Integer,default=0)
    size_22=db.Column(db.Integer,default=0)
    size_23=db.Column(db.Integer,default=0)
    size_24=db.Column(db.Integer,default=0)
    size_25=db.Column(db.Integer,default=0)
    size_26=db.Column(db.Integer,default=0)
    size_27=db.Column(db.Integer,default=0)
    size_28=db.Column(db.Integer,default=0)
    size_29=db.Column(db.Integer,default=0)
    size_30=db.Column(db.Integer,default=0)

    def __repr__(self):
        return('<Actual receive:packing_id - %s, colour - %s , total_qty -%spcs>'%(self.packing_id,self.colour,self.total_qty))


class Detail_carton(db.Model):
    __tablename__='detail_carton'
    id=db.Column(db.Integer,primary_key=True)
    packing_id=db.Column(db.Integer,db.ForeignKey('packings.id'))
    from_carton=db.Column(db.Integer,default=0)
    to_carton=db.Column(db.Integer,default=0)
    carton_qty=db.Column(db.Integer,default=0)
    colour=db.Column(db.String(20))
    per_carton_pcs=db.Column(db.Integer,default=0)
    per_carton_gw=db.Column(db.Integer,default=0)
    per_carton_nw=db.Column(db.Integer,default=0)
    subtotal=db.Column(db.Integer,default=0)
    length=db.Column(db.Float,default=0.0)
    width=db.Column(db.Float,default=0.0)
    height=db.Column(db.Float,default=0.0)
    size_1=db.Column(db.Integer,default=0)
    size_2=db.Column(db.Integer,default=0)
    size_3=db.Column(db.Integer,default=0)
    size_4=db.Column(db.Integer,default=0)
    size_5=db.Column(db.Integer,default=0)
    size_6=db.Column(db.Integer,default=0)
    size_7=db.Column(db.Integer,default=0)
    size_8=db.Column(db.Integer,default=0)
    size_9=db.Column(db.Integer,default=0)
    size_10=db.Column(db.Integer,default=0)
    size_11=db.Column(db.Integer,default=0)
    size_12=db.Column(db.Integer,default=0)
    size_13=db.Column(db.Integer,default=0)
    size_14=db.Column(db.Integer,default=0)
    size_15=db.Column(db.Integer,default=0)
    size_16=db.Column(db.Integer,default=0)
    size_17=db.Column(db.Integer,default=0)
    size_18=db.Column(db.Integer,default=0)
    size_19=db.Column(db.Integer,default=0)
    size_20=db.Column(db.Integer,default=0)
    size_21=db.Column(db.Integer,default=0)
    size_22=db.Column(db.Integer,default=0)
    size_23=db.Column(db.Integer,default=0)
    size_24=db.Column(db.Integer,default=0)
    size_25=db.Column(db.Integer,default=0)
    size_26=db.Column(db.Integer,default=0)
    size_27=db.Column(db.Integer,default=0)
    size_28=db.Column(db.Integer,default=0)
    size_29=db.Column(db.Integer,default=0)
    size_30=db.Column(db.Integer,default=0)
    size_31=db.Column(db.Integer,default=0)

    def __repr__(self):
        return('<Carton detail: packing_id - %s colour - %s, from %s to %s>'\
               %(self.packing_id,self.colour,self.from_carton,self.to_carton))



def make_shell_context():
    return dict(app=app,db=db,Order_info=Order_info)

manager.add_command("shell",Shell(make_context=make_shell_context))

if __name__=='__main__':
    manager.run()
