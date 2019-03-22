//Si estamos en un formulario de edicion, y dentro hay un combo que tira de Sharepoint, al entrar, se debe hacer una llamada, que pueda recuperar el ComboItemDTO que corresponde con el ID que permite elegir una option u otra.


//SHAREPOINT SERVICE!


/// Conociendo el ID de Sharepoint de una Persona, encuentra el ComboItemDTO de esa persona para pasarsela al combo
//NOTESE QUE LA VARIABLE DE ENTRADA ES EL ID (TIPO DE VARIABLE: INT/LONG)
public ComboItemDTO SelectStaffOption(long IDEmployeeSP)
{

	int NumberOfRows = 0;
	List<ListItem> ListaFinal = new List<ListItem>();
	List<ComboItemDTO> Combo2 = new List<ComboItemDTO>();
	ComboItemDTO ComboFiltered2 = new ComboItemDTO();
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
		int LastItemID = 0;
		for (int i = 0; i < IntegerNumber; i++)
		{
			if (i != 0)
			{
				LastItemID = (4900 * i) + 1;
			}
			List<ListItem> ListaIteraciones = new List<ListItem>();
			ListaIteraciones = (List<ListItem>)_SharepointQueryRepository.GetFragmentedStaffListByIDEmployee(IDEmployeeSP, LastItemID);
			 
			if (ListaIteraciones.Count > 0)
			{
				foreach (ListItem item in ListaIteraciones)
				{
					ListaFinal.Add(item);
				}
			}
		}

		foreach (ListItem Item in ListaFinal)
		{
			string Name = Item["Nombre"] != null ? (Item["Nombre"]).ToString() ?? "" : "";
			string Date = Item["FechaBaja"] != null ? (Item["FechaBaja"]).ToString() : String.Empty;
			string Surname = Item["Apellidos"] != null ? (Item["Apellidos"]).ToString() ?? "" : "";

			ComboItemDTO ComboItem = new ComboItemDTO
			{
				Value = (Item["ID"] != null) ? (int)Item["ID"] : 0,
				Text = string.Format("{0}, {1}", Surname, Name)
			};

			Combo2.Add(ComboItem);
		};

		foreach (ComboItemDTO item in Combo2)
		{
			if (item.Value == IDEmployeeSP)
			{
				ComboFiltered2 = item;
			}
		}

	}
	catch (Exception ex)
	{
		ExceptionHandling.ProcessException(ex, UtilesApplication.GetApplicationString("ActivePolicy"));
	}

	return ComboFiltered2;
}

//La funcion con nombre: GetFragmentedStaffListByIDEmployee, es la funcion que contiene la query que ejecuta sobre sharepoint

/// Devuelve las filas de personal cuyo titulo contenga algo parecido a la busqueda
public IList<ListItem> GetFragmentedStaffListByIDEmployee(long SearchIDEmployee, int LastItemId, SharepointConfigurationDTO SharepointConfig = null)
{
	IList<ListItem> ListItemObject = new List<ListItem>();

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
				ViewXml = @"<View><Query><Where><And><Gt><FieldRef Name='ID' /><Value Type='Counter'>" + LastItemId + "</Value></Gt><And><Eq><FieldRef Name='ID'/><Value Type='Counter'>" + SearchIDEmployee + "</Value></Eq>"
					 + "<IsNull><FieldRef Name='FechaBaja'/></IsNull>"
					 + "</And></And></Where></Query><OrderBy><FieldRef Name='ID' Ascending='True'/></OrderBy></View>"
			};

			ListItemCollection collListItem = oList.GetItems(Query);
			SharepointConfig.Context.Load(collListItem);
			SharepointConfig.Context.ExecuteQuery();

			foreach (ListItem oListItem in collListItem)
			{
				ListItemObject.Add(oListItem);
			}

		}
	}
	catch (Exception ex)
	{
		ExceptionHandling.ProcessException(ex, UtilesApplication.GetApplicationString("ActivePolicy"));
	}

	return ListItemObject;
}
