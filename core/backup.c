#include <stdio.h>
#include <stdlib.h>	//to use system()

int main()
{
	char *command = "streamlit run backup.py";
	
	if(system(command)){
        printf("Running command...\n");
    }
	return 0;
}