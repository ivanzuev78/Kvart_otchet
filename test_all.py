from Otchet_class import Cunductor, Shablon


def test_check_date_type():
    assert Cunductor.check_date_type('12-05-2021')
    assert not Cunductor.check_date_type('egewrrg')


def test_short_show():
    assert Cunductor.short_show([Shablon(0,0,0,'A70'), Shablon(0,0,0,'A71'), Shablon(0,0,0,'A72'), Shablon(0,0,0,'A73')]) == ['A70-73']
    assert Cunductor.short_show([Shablon(0,0,0,'A70'), Shablon(0,0,0,'A73'), Shablon(0,0,0,'A72'), Shablon(0,0,0,'A71')]) == ['A70-73']
    assert Cunductor.short_show([Shablon(0,0,0,'A80'), Shablon(0,0,0,'A73'), Shablon(0,0,0,'A72'), Shablon(0,0,0,'A71')]) == ['A71-73', 'A80']


def test_input_name():
    conductor = Cunductor()
    conductor.good_names['Имя2 И . И . '] = 'ДругеИмя И . О . '
    assert conductor.input_name('Имя И . О. ') == 'Имя И.О.'
    assert conductor.input_name('Имя2 И . И . ') == 'ДругеИмя И.О.'


def test_okonchanie():
    assert Cunductor.okonchanie('образец', 1) == '1 образец'
    assert Cunductor.okonchanie('образец', 2) == '2 образца'
    assert Cunductor.okonchanie('образец', 5) == '5 образцов'
    assert Cunductor.okonchanie('образец', 11) == '11 образцов'
