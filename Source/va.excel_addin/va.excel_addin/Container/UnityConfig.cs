namespace va.excel_addin_assignment.Container {
	using System;
	using Common;
	using Microsoft.Practices.Unity;

	public static class UnityConfig {
		/// <summary>
		/// Initate unitycontainer
		/// </summary>
		#region Unity Container
		private static Lazy<IUnityContainer> _container = new Lazy<IUnityContainer>(() =>
		{
			var containerInstance = new UnityContainer();
			RegisterTypes(containerInstance);
			return containerInstance;
		});

		/// <summary>
		/// Gets the configured Unity container.
		/// </summary>
		public static IUnityContainer GetConfiguredContainer() {
			return _container.Value;
		}
		#endregion
		public static void RegisterTypes(IUnityContainer container) {
			container.RegisterType(typeof(IDataLoader), typeof(DataLoader));
		}
	}
}
