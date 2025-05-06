Console.WriteLine("Preparing ...\n");

if (!Tools.TemplateFileInfo.Exists) {
	Console.WriteLine("模板文件不存在，请通过 RDP 软件创建 .rdp 文件，将 IP 地址修改为 \"${DefaultGateway}\"，\n并将 .rdp 文件重命名为 template.txt，放置在运行目录下。");
	return;
}

if (Tools.OutputFileInfo.Exists) {
	Console.WriteLine("检测到输出文件 output.rdp 已存在，为避免错误覆盖，请自行移动或删除输出文件。");
	return;
}

// 获取全部可用物理网络接口信息
var networkInterfaces = NetworkInterface.GetAllNetworkInterfaces()
	.Where(networkInterface =>
		networkInterface.OperationalStatus == OperationalStatus.Up
		&& networkInterface.Speed > 0
		&& (networkInterface.NetworkInterfaceType == NetworkInterfaceType.Ethernet || networkInterface.NetworkInterfaceType == NetworkInterfaceType.Wireless80211)
		&& !networkInterface.Description.Contains("Virtual")
		&& networkInterface.GetIPProperties().GatewayAddresses.Count > 0);

List<IPAddress> addresses = [];

// 展示所有可用的网络接口和 IPv4 默认网关地址
Console.WriteLine("找到以下网络接口和 IPv4 默认网关地址，请根据地址前的序号选择：");

foreach (var ni in networkInterfaces) {
	Console.WriteLine($"\n{ni.Name} ({ni.Description}) 本机 IPv4 地址：" +
		string.Join(", ", ni.GetIPProperties().UnicastAddresses.Where(ip => ip.Address.AddressFamily == AddressFamily.InterNetwork).Select(ip => ip.Address)));

	Console.WriteLine("IPv4 默认网关（请从中根据序号选择）：");
	foreach (var gatewayAddress in ni.GetIPProperties().GatewayAddresses.Where(g => g.Address.AddressFamily == AddressFamily.InterNetwork).Select(g => g.Address)) {
		Console.WriteLine($"[{addresses.Count}]  ".PadLeft(8) + gatewayAddress);
		addresses.Add(gatewayAddress);
	}
}

if (addresses.Count == 0) {
	Console.WriteLine("没有找到可用的 IPv4 默认网关地址，请检查网络连接或尝试重启计算机。");
	return;
}

var index = -1;

if (addresses.Count == 1) {
	index = 0;
	Console.WriteLine($"\n检测到有且仅有一个 IPv4 默认网关地址 ({addresses[index]})，将自动选择该地址。");
} else {
	// 选择默认网关
	Console.Write("\n请输入你需要选择的默认网关序号：");

	while (!int.TryParse(Console.ReadLine(), out index) || index < 0 || index >= addresses.Count) {
		Console.Write("输入错误，请重新输入：");
	}
}

var address = addresses[index];
Console.WriteLine($"你已成功选择 [{index}] {address}\n");

// 主动检查默认网关地址是否异常
if (address.ToString() is not { Length: >= 7 }) {
	Console.WriteLine("检测到 IPv4 默认网关地址异常，程序已终止，建议到 GitHub 存储库中发起议题。");
	return;
}

Console.WriteLine("Creating output file ...");

try {
	// 创建输出临时文件
	using var tempOutput = Tools.OutputTempFileInfo.CreateText();

	try {
		// 打开并读取模板文件
		using var template = Tools.TemplateFileInfo.OpenText();

		string? line = null;
		while ((line = template.ReadLine()) is not null) {
			// 尝试替换默认网关地址并写入输出文件
			tempOutput.WriteLine(line.Replace("${DefaultGateway}", address.ToString()));
		}

	} catch (Exception e) {
		Console.WriteLine($"读取模板文件失败或写入临时文件失败，异常信息：{e.Message}");

		try {
			// 显式确保释放资源，随后删除临时文件
			tempOutput.Close();
			tempOutput.Dispose();
			Tools.OutputTempFileInfo.Delete();
		} catch (Exception ex) {
			Console.WriteLine($"删除临时文件失败，建议手动移除 output.tmp，异常信息：{ex.Message}");
		}

		return;
	}

	try {
		// 显式确保写入并释放资源
		tempOutput.Flush();
		tempOutput.Close();
		tempOutput.Dispose();

		// 复制文件
		Tools.OutputTempFileInfo.CopyTo(Tools.OutputFileInfo.FullName, true);

		Console.WriteLine("输出文件创建完成，文件名为 output.rdp，请尽情使用。");
	} catch (Exception e) {
		Console.WriteLine($"复制输出文件失败，异常信息：{e.Message}");
	} finally {
		try {
			// 移除临时文件
			Tools.OutputTempFileInfo.Delete();
		} catch (Exception e) {
			Console.WriteLine($"其它操作均已完成，但删除临时文件失败，请手动移除 output.tmp，异常信息：{e.Message}");
		}
	}
} catch (Exception e) {
	Console.WriteLine($"创建输出文件失败，异常信息：{e.Message}");
	return;
}