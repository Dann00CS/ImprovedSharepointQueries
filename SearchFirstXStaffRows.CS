//Si estamos ante cualquier formulario, ya sea de creacion o de edicion, y tenemos un combo que tira de Sharepoint, el cual tiene un search input,
//cada vez que introduzcamos ahi algo, se debe buscar en Sharepoint, en busca de un elemento que contenga las letras escritas

/// Busca los primeros X Registros Activos de la lista de Personal para combo (Value: ID -- Text: "Title")
public List<ComboItemDTO> SearchFirstXStaffRows(string SearchInput)
{
	int NumberOfRows = 0;
	List<ComboItemDTO> Combo2 = new List<ComboItemDTO>();
	List<ComboItemDTO> ComboFiltered2 = new List<ComboItemDTO>();
	List<ListItem> SubListItems2 = new List<ListItem>();

	try
	{ 

		NumberOfRows = _SharepointQueryRepository.GetPersonalListSP_Count();
		double DecimalNumber = (double)NumberOfRows / 4900;
		int IntegerNumber = NumberOfRows / 4900;
		double resta = (double)DecimalNumber - IntegerNumber;
		if (resta > 0)
		{
			IntegerNumber++;
		}
		int MinimumItemId = 0;
		int MaximumItemId = 0;
		for (int i = 0; i < IntegerNumber; i++)
		{
			MinimumItemId = 4900 * (i);
			MaximumItemId = (4899 * (i + 1));
			if (MaximumItemId > NumberOfRows)
			{
				MaximumItemId = NumberOfRows - 1;
			}
			List<ListItem> ListaIteraciones = new List<ListItem>();
			ListaIteraciones = (List<ListItem>)_SharepointQueryRepository.GetFragmentedStaffListByNameEmployee(SearchInput, MinimumItemId, MaximumItemId);

			if (ListaIteraciones.Count > 0)
			{
				foreach (ListItem item in ListaIteraciones)
				{
					SubListItems2.Add(item);
				}
			}
		}

		foreach (ListItem Item in SubListItems2)
		{
			string Date = Item["FechaBaja"] != null ? (Item["FechaBaja"]).ToString() : String.Empty;
			string fullname = Item["Title"] != null ? (Item["Title"]).ToString() ?? "" : "";
			ComboItemDTO ComboItem = new ComboItemDTO
			{
				Value = (Item["ID"] != null) ? (int)Item["ID"] : 0,
				Text = string.Format("{0}", fullname)
			};

			Combo2.Add(ComboItem);
		};

		ComboFiltered2 = Combo2.Where(per => per.Text.ToLower().Contains(SearchInput.ToLower())).ToList();
	}
	catch (Exception ex)
	{
		ExceptionHandling.ProcessException(ex, UtilesApplication.GetApplicationString("ActivePolicy"));
	}

	return ComboFiltered2;
}

//La funcion con nombre: GetFragmentedStaffListByNameEmployee es la funcion que contiene la query que ejecuta sobre sharepoint

/// Busca las primeras X filas de Personal segun parametro NumberOfRows
public IList<ListItem> GetFragmentedStaffListByNameEmployee(string SearchNameEmployee, int MinimumItemId, int MaximumItemId, SharepointConfigurationDTO SharepointConfig = null)
{
	IList<ListItem> ListItemsCollection = new List<ListItem>();

	try
	{
		if (SharepointConfig == null)
		{
			SharepointConfig = GetSharepointConfig();
		}

		using (SharepointConfig.Context)
		{

			SP.List oList = SharepointConfig.Context.Web.Lists.GetByTitle("Personal");

			CamlQuery Query = new CamlQuery
			{
				ViewXml = @"<View><Query><Where><And>"
					 + "<Gt><FieldRef Name='ID' /><Value Type='Counter'>" + MinimumItemId + "</Value></Gt><And>" 
					 + "<Lt><FieldRef Name='ID' /><Value Type='Counter'>" + MaximumItemId + "</Value></Lt>" 
					 + "<And><Contains><FieldRef Name='Title'/><Value Type='Text'>" + SearchNameEmployee + "</Value></Contains>"
					 + "<IsNull><FieldRef Name='FechaBaja'/></IsNull>"
					 + "</And></And></And></Where></Query><OrderBy><FieldRef Name='ID' Ascending='True'/></OrderBy></View>"
			};

			ListItemCollection collListItem = oList.GetItems(Query);
			SharepointConfig.Context.Load(collListItem);
			SharepointConfig.Context.ExecuteQuery();

			foreach (ListItem oListItem in collListItem)
			{
				ListItemsCollection.Add(oListItem);
			}
		}
	}
	catch (Exception ex)
	{
		ExceptionHandling.ProcessException(ex, UtilesApplication.GetApplicationString("ActivePolicy"));
	}


	return ListItemsCollection;
}