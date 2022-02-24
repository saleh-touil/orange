#include <stdio.h>
#include <stdlib.h>	//to use system()

int main()
{
	char *command = "streamlit run app.py";
	printf("Running command...\n");
	system(command);
    while (1==1)
    {
    printf("");
    }
	return 0;
}