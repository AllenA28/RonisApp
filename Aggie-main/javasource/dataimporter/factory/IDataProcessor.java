package dataimporter.factory;

import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.systemwideinterfaces.core.IEntityProxy;
import com.mendix.systemwideinterfaces.core.IMendixObject;
import dataimporter.implementation.model.ImportMappingParameters;
import dataimporter.proxies.ColumnAttributeMapping;
import dataimporter.proxies.DataImporterElement;

import java.io.File;
import java.util.List;
import java.util.Map;

public interface IDataProcessor {
    Map<IEntityProxy, List<ColumnAttributeMapping>> startImport(IContext context, IMendixObject mappingTemplate);

    List<IMendixObject> parseData(IContext context, File file, IEntityProxy sheetMendixObject, List<ColumnAttributeMapping> columnAttributeMappingMendixObjects);

    Map<IEntityProxy, List<DataImporterElement>> startJsonMappingImport(IContext context, IMendixObject mappingTemplate);

    List<IMendixObject> parseJsonMappingData(IContext context, File file, IEntityProxy sheetMendixObject, List<DataImporterElement> dataImporterElementsMendixObjects, ImportMappingParameters mappingParameters);
}
