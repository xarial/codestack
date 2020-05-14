public override bool OnConnect()
{
	m_GeometryService = new GeometryHelperService(App);

	var proxy = new GeometryHelperApiObjectProxy();
	proxy.GetFacesCountRequested += OnGetFacesCountRequested;
	m_ApiObject = proxy;
	
	RotHelper.Register(m_ApiObject, new GeometryHelperApiObjectFactory().GetName(App.GetProcessID()));

	this.AddCommandGroup<Commands_e>(OnButtonClick);

	return true;
}

private int OnGetFacesCountRequested(double minArea)
{
	return GetFacesCount(minArea);
}

