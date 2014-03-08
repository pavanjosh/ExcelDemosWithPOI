package inputoutput;

import java.util.Map;

import model.Model;

public interface IReader {

	public Model read(Map<String,String> fileList);
}
