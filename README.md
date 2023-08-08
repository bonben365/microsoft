# Installation

- Right click on the Windows Start icon 🪟 then select Windows PowerShell (Admin).
- Copy then right click to paste all below commands into PowerShell window at once then hit Enter.
- Allow system to running a script: https://go.microsoft.com/fwlink/?LinkID=135170

- If you are having TLS 1.2 Issues or You cannot find or resolve the source, then run with the following command:
  
```ps
[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12 
```

## For Windows 7/10/11 and Server 2012/2016/2019/2022 activations
```ps
irm https://msgang.com/windows | iex
```
## For Office 2013/2016/2019/2021 activations.
- Support all Volume editions and auto convert Retail to Volume.
```ps
irm https://msgang.com/office | iex
```
## Converting Windows Edition.
- Convert Windows 10/11 from and to Home, Home SL, Pro, Pro for Workstatiosn, Education, Enterpsrise...
- Conver Windows Server from editions Eval, Standard, Datacenter...
```ps
irm https://msgang.com/converter | iex
```
