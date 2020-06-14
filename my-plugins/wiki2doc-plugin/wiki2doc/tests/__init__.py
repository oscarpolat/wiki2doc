from unittest import TestSuite, makeSuite

def test_suite():
    suite = TestSuite()
    from wiki2doc.tests import api, wiki2doc
    suite.addTest(makeSuite(api.Wiki2DocApiTestCase))
#    suite.addTest(makeSuite(api.AutoRepPermissionPolicyTestCase))
    suite.addTest(makeSuite(wiki2doc.Wiki2docTestCase))
    return suite
