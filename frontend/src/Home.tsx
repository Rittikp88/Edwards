import {
    useMemo,
    useCallback,
    useState,
    forwardRef,
    useImperativeHandle,
    useRef,
    useEffect,
    SetStateAction,
  } from "react";
  import "./App.css";
  import {
    Box,
    Text,
    Grid,
    GridItem,
    Breadcrumb,
    BreadcrumbItem,
    Button,
    BreadcrumbLink,
  } from "@chakra-ui/react";
  import { dummyData } from "./dummyData";
  import {
    Card,
    CardHeader,
    CardBody,
    CardFooter,
    Stack,
    Heading,
    StackDivider,
  } from "@chakra-ui/react";
  import { Input } from "@chakra-ui/react";
  import { GoArrowLeft } from "react-icons/go";
  import { useToast } from '@chakra-ui/react';
  import { BrowserRouter as Router, Route, Link ,Routes } from "react-router-dom";
  import Summary from "./Summary";
  
  interface Material {
    id: string;
    name: string;
    type: "material";
    studies: string[];
  }
  
  interface Study {
    id: string;
    name: string;
    type: "study";
    protocols: string[];
  }
  
  interface Protocol {
    id: string;
    name: string;
    type: "protocol";
    tests: string[];
  }
  
  interface Test {
    id: string;
    name: string;
    type: "test";
    duration: number;
  }
  
  interface resources {
    id: string;
    name: string;
    status: string;
  }
  
  export interface Data {
    materials: Material[];
    studies: Study[];
    protocols: Protocol[];
    tests: Test[];
    resources: resources[];
  }
  
  interface DirStructure<T> {
    id: string;
    name: string;
    data: T;
    type: "dir" | "file";
    children: DirStructure<T>[];
  }
  
  interface MaterialCardProps {
    // material: Material;
    onSelectMaterial: (material: Material) => void;
  }
  
  const MaterialCard = ({ onSelectMaterial,dummyDatas }:any) => {
    
    const dummy = dummyDatas[0]
    // console.log(dummy.materials,"bvfdvfhdgvfhdg")
    const [searchTerm, setSearchTerm] = useState("");
    const [selectedMaterialId, setSelectedMaterialId] = useState<number | null>(
      null
    );
    const [selectedMaterialIds, setSelectedMaterialIds] = useState<string[]>([]);
    const [selectedMaterial, setSelectedMaterial] = useState<string[]>([]);
    const [selectedStudies, setSelectedStudies] = useState<string[]>([]);
  
    const filteredMaterials = dummy.materials.filter((m:any) =>
      m.name.toLowerCase().includes(searchTerm.toLowerCase())
    );
  
    const toggleSelectMaterial = (materialId: any,materialName: any) => {
      setSelectedMaterialId(materialId);
      let materialIds = selectedMaterialIds
      // Toggle the selected state
      if (selectedMaterialIds&&selectedMaterialIds.includes(materialId)) {
        materialIds = selectedMaterialIds.filter((id) => id !== materialId)
        setSelectedMaterialIds(materialIds);
      } else {
        materialIds = [...selectedMaterialIds, materialId]
        setSelectedMaterialIds(materialIds);
      }
       // Remove duplicates from studies and update selectedStudies
       const allStudies = materialIds
       .map((id) => dummy.materials.find((material:any) => material.id === id)?.studies || [])
       .flat();
       console.log(allStudies,"allStudies......")
     const uniqueStudies = [...new Set(allStudies)];
     setSelectedStudies(uniqueStudies);
     onSelectMaterial(uniqueStudies,materialName)
  
  // console.log(uniqueStudies,"selectedStudies.......")
      
    };     
    return (
      <Card
        style={{ backgroundColor: "rgba(255, 0, 0, 0.6)", padding: "10px" }}
        maxW="250px"
        w="100%"
        p="10px"
        bgGradient={[
          'linear(to-tr, teal.300, yellow.400)',
          'linear(to-t, blue.200, teal.500)',
          'linear(to-b, orange.100, purple.300)',
        ]}
      >
        <CardHeader display="flex" justifyContent="space-between">
          <Heading size="md" marginTop={"-4px"}>Materials</Heading>
          <Input
            htmlSize={7}
            width="90px"
            ml="30px"
            size="sm"
            placeholder="search"
            variant="filled"
            onChange={(e) => setSearchTerm(e.target.value)}
            marginTop={"-4px"}
          />
        </CardHeader>
  
        <CardBody>
          <Stack divider={<StackDivider />} spacing="4">
            {filteredMaterials.map((material: any) => (
              <Box
                key={material.id}
                p="4"
                // borderWidth="1px"
                // borderRadius="md"
                _hover={{
                  bg: "green.500", // Background color on hover>
                  // boxShadow: "0 0 10px rgba(255, 255, 255, 0.8)",
                }}
                onClick={() => {
                  toggleSelectMaterial(material.id,material.name);
                  // setSelectedMaterialId(material.id);
                  setSelectedMaterial(material.name)
                }}
                bg={"white"}
                bgGradient={selectedMaterialIds.includes(material.id) ? 'linear(to-r, teal.500, green.500)' : "white"}
              >
                <Heading size="sm" color={selectedMaterialIds.includes(material.id) ? "white" : "black"}>{material.name}</Heading>
              </Box>
            ))}
          </Stack>
        </CardBody>
      </Card>
    );
  };
  
  const StudyCard = ({ selectedMaterial, onSelectStudies }: any) => {
    const [searchTerm, setSearchTerm] = useState("");
    const [selectedMaterialId, setSelectedMaterialId] = useState<number | null>(
      null
    );
    const [selectedMaterialIds, setSelectedMaterialIds] = useState<string[]>([]);
    const [selectedStudies, setSelectedStudies] = useState<string[]>([]);
    console.log("studies", selectedMaterial.studies);
  
    const filteredStudies = dummyData.studies.filter(
      (study:any) =>
      selectedMaterial&&selectedMaterial.includes(study.id) &&
        study.name.toLowerCase().includes(searchTerm.toLowerCase())
    );
  
    const toggleSelectStudies = (materialId: any, materialName: any) => {
      setSelectedMaterialId(materialId);
      let materialIds = selectedMaterialIds
      // Toggle the selected state
      if (selectedMaterialIds&&selectedMaterialIds.includes(materialId)) {
        materialIds = selectedMaterialIds.filter((id) => id !== materialId)
        setSelectedMaterialIds(materialIds);
      } else {
        materialIds = [...selectedMaterialIds, materialId]
        setSelectedMaterialIds(materialIds);
      }
       // Remove duplicates from studies and update selectedStudies
       const allStudies = materialIds
       .map((id) => dummyData.studies.find((material:any) => material.id === id)?.protocols || [])
       .flat();
       console.log(allStudies,"allStudies......")
     const uniqueStudies = [...new Set(allStudies)];
     setSelectedStudies(uniqueStudies);
     onSelectStudies(uniqueStudies,materialName)
  
  // console.log(uniqueStudies,"selectedStudies.......")
    }
    const handleArrow = () => {
      onSelectStudies(null)
    }
    return (
      <Card
        style={{ backgroundColor: "rgba(255, 0, 0, 0.6)" }}
        maxW="250px"
        w="100%"
        bgGradient={[
          'linear(to-tr, teal.300, yellow.400)',
          'linear(to-t, blue.200, teal.500)',
          'linear(to-b, orange.100, purple.300)',
        ]}
      >
        <CardHeader display="flex" justifyContent="space-between">
          <GoArrowLeft size={30} style={{ cursor: "pointer" }} onClick={handleArrow}/>
          <Heading size="md">Studies</Heading>
          <Input
            htmlSize={7}
            width="90px"
            ml="30px"
            size="sm"
            placeholder="search"
            variant="filled"
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </CardHeader>
  
        <CardBody>
          <Stack divider={<StackDivider />} spacing="4">
            {filteredStudies.map((material: any) => (
              <Box
                key={material.id}
                p="4"
                // borderWidth="1px"
                // borderRadius="md"
                _hover={{
                  bg: "green.500", // Background color on hover>
                }}
                onClick={() => {
                  setSelectedMaterialId(material.id);
                  toggleSelectStudies(material.id,material.name);
                }}
                bg={"white"}
                bgGradient={selectedMaterialIds.includes(material.id) ? 'linear(to-r, teal.500, green.500)' : "white"}
              >
                <Heading size="sm" color={selectedMaterialIds.includes(material.id) ? "white" : "black"}>{material.name}</Heading>
              </Box>
            ))}
          </Stack>
        </CardBody>
      </Card>
    );
  };
  
  const Protocols = ({ selectedStudies, onSelectProtocols }: any) => {
    const [searchTerm, setSearchTerm] = useState("");
    const [selectedMaterialId, setSelectedMaterialId] = useState<number | null>(
      null
    );
    const [selectedMaterialIds, setSelectedMaterialIds] = useState<string[]>([]);
    const [selectedTest, setSelectedStudies] = useState<string[]>([]);
    // console.log("studies", selectedStudies.protocols);
  
    const filteredStudies = dummyData.protocols.filter(
      (study:any) =>
        selectedStudies.includes(study.id) &&
        study.name.toLowerCase().includes(searchTerm.toLowerCase())
    );
  
    const toggleSelectProtocols = (materialId: any, materialName:any) => {
      setSelectedMaterialId(materialId);
      let materialIds = selectedMaterialIds
      // Toggle the selected state
      if (selectedMaterialIds&&selectedMaterialIds.includes(materialId)) {
        materialIds = selectedMaterialIds.filter((id) => id !== materialId)
        setSelectedMaterialIds(materialIds);
      } else {
        materialIds = [...selectedMaterialIds, materialId]
        setSelectedMaterialIds(materialIds);
      }
       // Remove duplicates from studies and update selectedStudies
       const allStudies = materialIds
       .map((id) => dummyData.protocols.find((material:any) => material.id === id)?.tests || [])
       .flat();
      //  console.log(allStudies,"allStudies......")
     const uniqueStudies = [...new Set(allStudies)];
     setSelectedStudies(uniqueStudies);
     onSelectProtocols(uniqueStudies,materialName)
  
  // console.log(uniqueStudies,"selectedStudies.......")
  
    }
    const handleArrow = () => {
      onSelectProtocols(null)
    }
    return (
      <Card
        style={{ backgroundColor: "rgba(255, 0, 0, 0.6)" }}
        maxW="250px"
        w="100%"
        bgGradient={[
          'linear(to-tr, teal.300, yellow.400)',
          'linear(to-t, blue.200, teal.500)',
          'linear(to-b, orange.100, purple.300)',
        ]}
      >
        <CardHeader display="flex" justifyContent="space-between">
          <GoArrowLeft size={30} style={{ cursor: "pointer" }} onClick={handleArrow} />
          <Heading size="md">Protocols</Heading>
          <Input
            htmlSize={7}
            width="90px"
            ml="30px"
            size="sm"
            placeholder="search"
            variant="filled"
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </CardHeader>
  
        <CardBody>
          <Stack divider={<StackDivider />} spacing="4">
            {filteredStudies.map((material: any) => (
              <Box
                key={material.id}
                p="4"
                // borderWidth="1px"
                // borderRadius="md"
                _hover={{
                  bg: "green.500", // Background color on hover>
                }}
                onClick={() => {
                  // onSelectProtocols(material);
                  toggleSelectProtocols(material.id, material.name)
                  setSelectedMaterialId(material.id);
                }}
                bg={"white"}
                bgGradient={selectedMaterialIds.includes(material.id) ? 'linear(to-r, teal.500, green.500)' : "white"}
              >
                <Heading size="sm" color={selectedMaterialIds.includes(material.id) ? "white" : "black"}>{material.name}</Heading>
              </Box>
            ))}
          </Stack>
        </CardBody>
      </Card>
    );
  };
  
  const Test = ({ selectedProtocols, onSelectTest, selectedResource }: any) => {
    const [searchTerm, setSearchTerm] = useState("");
    const [selectedMaterialId, setSelectedMaterialId] = useState<number | null>(
      null
    );
  
    console.log("studies", selectedProtocols.tests);
  
    const filteredStudies = dummyData.tests.filter(
      (study:any) =>
        selectedProtocols.includes(study.id) &&
        study.name.toLowerCase().includes(searchTerm.toLowerCase())
    );
  
    const handleArrow = () => {
      onSelectTest(null)
    }
    return (
      <Card
        style={{ backgroundColor: "rgba(255, 0, 0, 0.6)" }}
        maxW="250px"
        w="100%"
        bgGradient={[
          'linear(to-tr, teal.300, yellow.400)',
          'linear(to-t, blue.200, teal.500)',
          'linear(to-b, orange.100, purple.300)',
        ]}
      >
        <CardHeader display="flex" justifyContent="space-between">
          <GoArrowLeft size={30} style={{ cursor: "pointer" }} onClick={handleArrow}/>
          <Heading size="md">Test</Heading>
          <Input
            htmlSize={7}
            width="90px"
            ml="30px"
            size="sm"
            placeholder="search"
            variant="filled"
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </CardHeader>
  
        <CardBody>
          <Stack divider={<StackDivider />} spacing="4">
            {filteredStudies.map((material: any) => (
              <Box
                key={material.id}
                p="4"
                // borderWidth="1px"
                // borderRadius="md"
                _hover={{
                  bg: "green.500", // Background color on hover>
                }}
                onClick={() => {
                  // toggleSelectTest(material.name);
                  // onSelectProtocols(material);
                  onSelectTest(material,material.name);
                  setSelectedMaterialId(material.id);
                }}
                bg={"white"}
                bgGradient={selectedMaterialId === material.id ? 'linear(to-r, teal.500, green.500)' : "white"} // Background color if selected
              >
                <Heading size="sm" color={selectedMaterialId === material.id? "white" : "black"}>{material.name}</Heading>
                <Text>Duration: {material.duration}d</Text>
              </Box>
            ))}
          </Stack>
        </CardBody>
      </Card>
    );
  };
  
  const Resources = ({ onSelectResources, selectedTest ,selMaterial,selStudies,selProtocols}: any) => {
    const {name, duration} = selectedTest;
    const [resources, setResources] = useState([]);
    const [loading, setLoading] = useState(true);
    const [selectedMaterialId, setSelectedMaterialId] = useState<number | null>(
      null
    );
    const [searchTerm, setSearchTerm] = useState("");
    const [selectResources, setSelectResources] = useState({});
    const [dataChanged, setDataChanged] = useState(false);
    const toast = useToast()
    // console.log("svdhachdfgascfga");
    useEffect(() => {
      const fetchData = async () => {
        try {
          const response = await fetch(
            "https://6596915d6bb4ec36ca02eba3.mockapi.io/resource"
          );
          const data = await response.json();
          setResources(data);
          setLoading(false);
        } catch (error) {
          console.error("Error fetching data:", error);
          setLoading(false);
        }
      };
      fetchData();
    }, [dataChanged]);
    const filteredStudies = resources.filter(
      (study:any) =>
        study.name.toLowerCase().includes(searchTerm.toLowerCase())&&
        study.status !== "Leave"
    );
  
    const handleAllocate = async (selectedMaterialId:any) => {
      const generateRandomNumber = () => {
        return Math.floor(Math.random() * 9000) + 1000;
      };
      
      
        const Eqlnumber = `EQL-${generateRandomNumber()}`;
    
      console.log(selectedMaterialId)
      try {
        // Perform a PUT request to update the resource
        const response = await fetch(
          `https://6596915d6bb4ec36ca02eba3.mockapi.io/resource/${selectedMaterialId}`,
          {
            method: "PUT",
            headers: {
              "Content-Type": "application/json",
            },
            body: JSON.stringify({
              ...selectResources,
              test: name, // Update the 'test' field with the new value
              duration: `${duration}`,
              status: "Allocated",
              Materials: selMaterial,
              Studies: selStudies,
              Protocols: selProtocols,
              EQL:Eqlnumber
              
            }),
          }
        );
  
        if (!response.ok) {
          throw new Error(`Failed to update resource: ${response.statusText}`);
        }
        setDataChanged(true)
         // Reload the window after 4 seconds

     setTimeout(() => {
      // Reload the window
window.location.reload();


     },5000)
  
        // Update the local state with the updated resource
        // setResources((prevResources) =>
        //   prevResources.map((r) => (r.id === resource.id ? { ...r, test: 'UpdatedTest' } : r))
        // );
      } catch (error) {
        console.error("Error updating resource:", error);
      }
    };
    if (loading) {
      return <div>Loading...</div>;
    }
  
    return (
      <div style={{display: "flex", justifyContent: "space-between"}}>
        {filteredStudies.length === 0 ? (
        <Card  maxW="250px" w="100%"  bgGradient={[
          'linear(to-tr, teal.300, yellow.400)',
          'linear(to-t, blue.200, teal.500)',
          'linear(to-b, orange.100, purple.300)',
        ]}>
          <CardHeader display="flex" justifyContent="space-between">
            <Heading size="md">Resources</Heading>
            <Input
              htmlSize={7}
              width="90px"
              ml="30px"
              size="sm"
              placeholder="search"
              variant="filled"
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </CardHeader>
          <CardBody>
            <Text>No data available for resources.</Text>
          </CardBody>
        </Card>
      ) : (
      <Card
        style={{ backgroundColor: "rgba(255, 0, 0, 0.6)" }}
        maxW="250px"
        w="100%"
        bgGradient={[
          'linear(to-tr, teal.300, yellow.400)',
          'linear(to-t, blue.200, teal.500)',
          'linear(to-b, orange.100, purple.300)',
        ]}
      >
        <CardHeader display="flex" justifyContent="space-between">
          {/* <GoArrowLeft size={30} style={{ cursor: "pointer" }} /> */}
          <Heading size="md">Resources</Heading>
          <Input
            htmlSize={7}
            width="90px"
            ml="30px"
            size="sm"
            placeholder="search"
            variant="filled"
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </CardHeader>
  
        <CardBody>
          <Stack divider={<StackDivider />} spacing="4">
            {filteredStudies.map((material: any) => (
              <Box
                key={material.id}
                p="4"
                // borderWidth="1px"
                // borderRadius="md"
                _hover={{
                  bg: "green.500", // Background color on hover>
                }}
                onClick={() => {
                  // onSelectProtocols(material);
                  setSelectedMaterialId(material.id);
                  setSelectResources(material);
                }}
                bg={"white"}
                bgGradient={selectedMaterialId === material.id ? 'linear(to-r, teal.500, green.500)' : "white"} // Background color if selected
              >
                <div style={{ display: "flex", justifyContent: "space-between" }}>
                  <Heading size="sm" color={selectedMaterialId === material.id? "white" : "black"}>{material.name}</Heading>
                  {material.status == "Allocated" &&<Button
                  size='xs'
                    color={
                      material.status === "Available" ? "green.500" : "red.500"
                    }
                    borderRadius="md"
                    p="1"
                    bg={material.status === "Available" ? "green.100" : "red.100"}
                  >
                    Allocated-{material.duration}d
                  </Button>}
                  {/* {<Button size='xs' onClick={() => handleAllocate(selectedMaterialId)}>Allocate</Button>} */}
                  {material.status == "Available" && <Button
                  // onClick={() => handleAllocate(selectedMaterialId)}
                  size='xs'
                    color={
                      material.status === "Available" ? "green.500" : "red.500"
                    }
                    borderRadius="md"
                    p="1"
                    bg={material.status === "Available" ? "green.100" : "red.100"}
                  >Available</Button>}
                </div>
                {material.test !== "Null"&&<Text>{material.test}</Text>}
                
              </Box>
            ))}
          </Stack>
        </CardBody>
      </Card>
      )}
      <Button style={{marginLeft: "40px"}} bgGradient='linear(to-r, teal.500, green.500)'
    _hover={{
      bgGradient: 'linear(to-r, red.500, yellow.500)',
    }} color={"white"} onClick={() => {
      handleAllocate(selectedMaterialId)
       // Create an example promise that resolves in 5s
       const examplePromise = new Promise((resolve, reject) => {
        setTimeout(() => resolve(200), 3000)
      })
  
      // Will display the loading toast until the promise is either resolved
      // or rejected.
      toast.promise(examplePromise, {
        success: { title: 'Allocated', description: 'Looks great' },
        error: { title: 'Promise rejected', description: 'Something wrong' },
        loading: { title: 'Pending', description: 'Please wait' },
      })
      }}>Allocate</Button>
      </div>
    );
  };
  
  function Home() {
    const [selectedMaterial, setSelectedMaterial] = useState(null);
    const [selMaterial, setSelMaterial] = useState(null);
    const [selectedStudies, setSelectedStudies] = useState(null);
    const [selStudies, setSelStudies] = useState(null);
    const [selectedProtocols, setSelectedProtocols] = useState(null);
    const [selProtocols, setSelProtocols] = useState(null);
    const [selectedTest, setSelectedTest] = useState(null);
    const [selTest, setSelTest] = useState(null);
    const [selectedResource, setSelectedResource] = useState(null);
    const [dummyDatas, setDummyData] = useState(null);
  useEffect(() => {
  const dummyData = async() => {
    try{
      const materialsResponse = await fetch('https://658d4ccc7c48dce94738dde6.mockapi.io/materials/materials');
      const materialsData = await materialsResponse.json();
      console.log(materialsData)
      setDummyData(materialsData)
  
  } catch (error) {
    console.error('Error fetching data:', error);
  }
  };
  dummyData();
  },[]);
  
  
    const onSelectMaterial = (material: SetStateAction<null>,selectedMaterial: SetStateAction<null>) => {
      setSelectedMaterial(material);
      console.log(selectedMaterial,"duidudidudi")
      setSelMaterial(selectedMaterial);
      setSelectedStudies(null);
      setSelectedProtocols(null);
      setSelectedTest(null);
    };
    const onSelectStudies = (material: any, materialName:any) => {
      setSelStudies(materialName)
      setSelectedStudies(material);
      setSelectedProtocols(null);
      setSelectedTest(null);
    };
  
    const onSelectProtocols = (material: any, materialName:any ) => {
      setSelProtocols(materialName);
      setSelectedProtocols(material);
      setSelectedTest(null);
      // console.log(selectedProtocols)
    };
    const onSelectTest = (material: any,materialName:any) => {
      setSelTest(materialName);
      setSelectedTest(material);
      // console.log(selectedTest)
    };
    const onSelectResources = (resource: any) => {
      setSelectedResource(resource);
      console.log(selectedResource);
    };
  
    return (
      <Box p={4}>
        <Box
    p={4}
    bgGradient="linear-gradient(to right, #7E5BEF, #4D39CE)"
    boxShadow="0px 2px 4px rgba(0, 0, 0, 0.1)"
    borderBottom="4px solid #592DEA"
  >
    <div
      style={{
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
      }}
    >
       <Text
      //  bg={}
            bgGradient="linear(to-l, #FF0080, #FF0080)"
            bgClip="text"
            fontSize="2xl"
            fontWeight="extrabold"
          >
            Edwards
          </Text>
      <Link to="/summary">
        <Button colorScheme="teal" variant="solid">
          Summary & Progress Report
        </Button>
      </Link>
    </div>
    
  </Box>
        <div style={{ display: "flex", width: "100%" }}>
          <div style={{ display: "flex", flexWrap: "wrap", justifyContent: 'space-between'}}>
            <div style={{ padding: "10px",flex: "1" }}>
              
              {dummyDatas && (<MaterialCard 
              onSelectMaterial={onSelectMaterial}
              dummyDatas = {dummyDatas}
               /> )}
            </div>
            <div style={{ padding: "10px",flex: "1"}}>
              {selectedMaterial && (
                <StudyCard
                  selectedMaterial={selectedMaterial}
                  onSelectStudies={onSelectStudies}
                />
              )}
            </div>
  
            <div style={{ padding: "10px",flex: "1" }}>
              {selectedStudies && (
                <Protocols
                  selectedStudies={selectedStudies}
                  onSelectProtocols={onSelectProtocols}
                />
              )}
            </div>
            <div style={{ padding: "10px",flex: "1" }}>
              {selectedProtocols && (
                <Test
                  selectedProtocols={selectedProtocols}
                  onSelectTest={onSelectTest}
                  selectedResource={selectedResource}
                />
              )}
            </div>
            <div style={{ padding: "10px",flex: "1" }}>
              {selectedTest && (
                <Resources
                  onSelectResources={onSelectResources}
                  selectedTest={selectedTest}
                  // allData={[selMaterial,selStudies,selProtocols,selTest]}
                  selMaterial = {selMaterial}
                  selStudies = {selStudies}
                  selProtocols= {selProtocols}
                />
              )}
            </div>
          </div>
          {/* <div style={{width: "20%", backgroundColor: "aqua"}}>
        <Resources 
        onSelectResources = {onSelectResources}
        />
        </div> */}
        </div>
      </Box>
    );
  }
  
  export default Home;
  