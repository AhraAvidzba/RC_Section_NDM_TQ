"""ДЛЯ ТЕСТОВ. Переиспользует независимую сборку посткоммитных напряжений для нового проверенного состояния20."""
from pathlib import Path
here=Path(__file__).resolve().parent
source=here/'Check-TestColdAcceptedStressContinuity.py';code=source.read_text('utf-8-sig')
code=code.replace("Cold862ColdAcceptedStressContinuity_20261125","Cold869FixedMetricNextScan_20261126").replace('Cold858BalancedInitialStressReplay_20261125','Cold867FixedMetricKernel_20261126').replace('Cold859SpatialDeclineStressReplay_20261125','Cold867FixedMetricKernel_20261126')
old="name='Native/EnvelopeEnergy.Accepted'+suffix;assert (root/name).read_bytes()==(verified/name).read_bytes(),name"
new="""name='Native/EnvelopeEnergy.Accepted'+suffix
    left=(root/name).read_bytes();right=(verified/('Native/LoadingWalk.Accepted.Sequence'+suffix)).read_bytes()
    if suffix=='.csv':
        a=left.splitlines();b=right.splitlines();assert len(a)==len(b)
        different=[i for i,(x,y) in enumerate(zip(a,b)) if x!=y]
        assert different==[0,6,7]
        assert a[0]==b'MODE;AcceptedPostCrackTestPath;NotACapacityResult' and b[0]==b'MODE;DiagnosticReSolve;NotAcceptedContinuation;NotACapacityResult'
        assert a[6].startswith(b'META;4;0;') and b[6].startswith(b'META;4;0;')
        assert a[7]==b'DETAILS;AcceptedPostCrackTestPath;capacity_certified=False' and b[7]==b'DETAILS;DiagnosticReSolve;capacity_certified=False'
        assert [x for i,x in enumerate(a) if i not in different]==[x for i,x in enumerate(b) if i not in different]
    else:assert left==right
"""
assert code.count(old)==1;code=code.replace(old,new)
assert code.count("len(source['fields'])==19")==1;code=code.replace("len(source['fields'])==19","len(source['fields'])==20")
assert code.count("int(row[1])==19")==1;code=code.replace("int(row[1])==19","int(row[1])==20")
code=code.replace('exact19_accepted_stresses_and_operators_equal_verified859','exact20_accepted_stresses_and_operators_equal_verified867')
code=code.replace('zero_increment_repeated24_times_per_point_is_bit_identical','zero_increment_repeated10_times_per_point_is_numerically_identical')
exec(compile(code,str(source),'exec'),{'__file__':str(source),'__name__':'verified20_postcommit_scan_test'})
