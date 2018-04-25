namespace va.excel_addin_assignment.Common {
	using System;
	using System.Collections.Generic;
	using System.Linq;
	using System.Text;
	using System.Threading.Tasks;
	using Container;
	using Microsoft.Practices.Unity;

	public static class AddInManager {

		public static void Initialize() {
			GetUnityInstance();
		}

		public static void ShutDown() {
			GetUnityInstance().Dispose();
		}
		public static IUnityContainer GetUnityInstance() {
			return UnityConfig.GetConfiguredContainer();
		}
	}
}
