Console.WriteLine("Preparing ...\n");

if (!Tools.TemplateFileInfo.Exists) {
	Console.WriteLine("模板文件不存在，请通过 RDP 软件创建 .rdp 文件，将 IP 地址修改为 \"${DefaultGateway}\"，\n并将 .rdp 文件重命名为 template.txt，放置在运行目录下。");
	return;
}

var isManualMode = args is ["manual"];

if (isManualMode) {
	Console.WriteLine("手动模式已启用。\n");
}

if (!isManualMode && Tools.OutputFileInfo.Exists) {
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
IPAddress? address = null;

if (!isManualMode && addresses.Count == 1) {
	index = 0;
	Console.WriteLine($"\n检测到有且仅有一个 IPv4 默认网关地址 ({addresses[index]})，将自动选择该地址。");
} else {
	// 选择默认网关
	Console.Write("\n请输入你需要选择的默认网关序号或直接输入任意完整 IPv4 地址：");

	while (true) {
		var input = Console.ReadLine();
		// 尝试解析为索引
		if (int.TryParse(input, out var parsedIndex) && parsedIndex >= 0 && parsedIndex < addresses.Count) {
			index = parsedIndex;
			break;
		}
		// 尝试解析为 IP 地址
		if (IPAddress.TryParse(input, out var parsedAddress) && parsedAddress.AddressFamily == AddressFamily.InterNetwork) {
			address = parsedAddress;
			break;
		}
		Console.Write("输入错误，请重新输入：");
	}
}

address ??= addresses[index];
Console.WriteLine($"你已成功选择 [{index}] {address}\n");

var port = 3389;

if (isManualMode) {
	// 修改端口号
	Console.Write("请输入远程桌面连接端口号（直接回车使用默认值 3389）：");
	var portInput = Console.ReadLine();
	if (string.IsNullOrWhiteSpace(portInput)) {
		// 使用默认端口号
	} else if (int.TryParse(portInput, out port) && port > 0 && port <= 65535) {
		Console.WriteLine($"已成功设置端口号为 {port}\n");
	} else {
		port = 3389;
		Console.WriteLine("输入无效，已使用默认端口号 3389\n");
	}
}

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
			tempOutput.WriteLine(line.Replace("${DefaultGateway}", port == 3389 ? address.ToString() : string.Concat(address, ":", port)));
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

		if (isManualMode && OperatingSystem.IsWindows()) {
			Console.Write("按 y 自动启动远程桌面连接，按其他任意键退出程序：");
			var key = Console.ReadKey();
			Console.WriteLine();
			if (key.KeyChar == 'y' || key.KeyChar == 'Y') {
				try {
					// 启动远程桌面连接
					using var process = new System.Diagnostics.Process();
					process.StartInfo.FileName = Path.Join(Environment.SystemDirectory, "mstsc.exe");
					process.StartInfo.Arguments = Tools.OutputFileInfo.FullName;
					process.Start();
					Console.WriteLine("远程桌面连接已启动，祝你使用愉快！");
				} catch (Exception e) {
					Console.WriteLine($"远程桌面连接启动失败，异常信息：{e.Message}");
				}
			}
		}
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