ALTER TABLE `peripheral` 
	CHANGE `pertype` `pertype` 
    enum('Bateria','Camara',
    'Cargador','Chip','Consola',
    'Defcom','Docking','Eq. Video Confer.',
    'Fax','Filmadora','GPS','Grab. Digital',
    'HDD','Lampara','Lector Cod.Barra','LectorDVD Ext.',
    'Maletin','Microfono','Modem','Monitor','Mouse',
    'Parlante','Pendrive','Plantroncs','Proyector',
    'Puntero','Ram','Scanner','TVs','Tarjeta_Video',
    'Tarjeta_Wifi','Teclado','Telon','Varios','DOCK'
    ,'Mochila','Cable','Adaptador','Funda','Auriculares'
    ,'Funcional','Memoria','Almacenamiento') CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci DEFAULT NULL ;