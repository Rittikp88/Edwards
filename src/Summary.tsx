import React from "react";
import {
  Box,
  Text,
  Image,
  Divider,
  Button,
  Grid,
  GridItem,
  Heading,
  Table,
  TableCaption,
  TableContainer,
  Tbody,
  Td,
  Tfoot,
  Th,
  Thead,
  Tr,
  Select,
  useDisclosure,
  Flex,
} from "@chakra-ui/react";
import axios from "axios";
import { Link } from "react-router-dom"; // Assuming you are using React Router for navigation
import { PiAddressBookFill } from "react-icons/pi";
import { BsGraphUpArrow } from "react-icons/bs";
import { FaNoteSticky } from "react-icons/fa6";
import { Chart } from "react-google-charts";
import { Badge } from "@chakra-ui/react";
import SummaryModel from "./SummaryModel";
import { dummyData } from "./data/dummydata";


// export const data = [
//   ["Availability", "Numbers"],
//   ["Available", 11],
//   ["leave", 2],
//   ["Allocated", 7],
// ];


export const options = {
  title: "Resource Availability",
};

const TableHeader = ({ selectedStatus, setSelectedStatus }: any) => (
  <Thead style={{ position: "sticky", top: 0, zIndex: 1, background: "white" }}>
    <Tr bg={"purple.500"}>
      <Th color={"white"}>EQL</Th>
      <Th color={"white"}>Name</Th>
      <Th color={"white"}>Skills</Th>
      <Th isNumeric color={"white"}>
        <Select
          // placeholder="Status"
          value={selectedStatus || "Status"}
          onChange={(e) => setSelectedStatus(e.target.value)}
          color={"black"}
        >
          <option value="Status">Status</option>
          <option value="Available">Available</option>
          <option value="Leave">Not Available</option>
          <option value="Allocated">Allocated</option>
        </Select>
      </Th>
      <Th color={"white"}>Report</Th>
    </Tr>
  </Thead>
);
const NavigationBar = () => (
  <Flex justifyContent="space-between" alignItems="center" mb={4}>
    <Text
            bgGradient="linear(to-l, #7928CA, #FF0080)"
            bgClip="text"
            fontSize="2xl"
            fontWeight="extrabold"
          >
            Edwards
          </Text>
    <Flex alignItems="center">
      <Link to="/" style={{ textDecoration: "none" }}>
        <Button ml={4} colorScheme="teal">
          Back
        </Button>
      </Link>
    </Flex>
  </Flex>
);


function Summary() {
  const [selectedStatus, setSelectedStatus] = React.useState("Status");
  // const { isOpen, onOpen, onClose } = useDisclosure()
  const [resources, setResources] = React.useState([]);
  // Assuming 'resources' is your fetched data array
const totalDataCount = resources.length;

const totalAllocatedCount = resources.filter((item : any) => item.status === "Allocated").length;
const totalAvailableCount = resources.filter((item: any) => item.status === "Available").length;
const totalLeaveCount = resources.filter((item: any) => item.status === "Leave").length;

const data = [
  ["Availability", "Numbers"],
  ["Available", totalAvailableCount],
  ["Leave", totalLeaveCount],
  ["Allocated", totalAllocatedCount],
];

  React.useEffect(() => {
    // Fetch data from the API
    const fetchData = async () => {
      try {
        const response = await axios.get(
          "https://6596915d6bb4ec36ca02eba3.mockapi.io/resource"
        );
        setResources(response.data);
      } catch (error) {
        console.error("Error fetching data:", error);
      }
    };

    fetchData();
  }, []); // Empty dependency array means this effect runs once after the initial render

  const handleDownloadCSV = () => {
    const availableData = resources.filter((item: any) => item.status === "Allocated");

    const csvContent =
      "data:text/csv;charset=utf-8," +
      "Id,Name,Duration,Skills,Status,test,Materials, Studies,Protocols, EQL\n" +
      availableData
        .map((item : any) =>
          [
            item.id,
            item.name,
            item.duration,
            item.skills,
            item.status,
            item.test,
            item.Materials,
            item.Studies,
            item.Protocols,
            item.EQL
          ].join(",")
        )
        .join("\n");

    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "summary-report.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const generateRandomNumber = () => {
    return Math.floor(Math.random() * 9000) + 1000;
  };
  
  const generateEQL = () => {
    return `EQL-${generateRandomNumber()}`;
  };

  const filteredData = (selectedStatus === "Status"
  ? resources
  : resources.filter((item: { status: string }) => item.status === selectedStatus)) as Array<{
    id: string;
    name: string;
    skills: string;
    status: string;
    EQL: string
    // Add other properties as needed
  }>;
      

  return (
    <>
      {/*  */}
      <Box
        w="100%"
        h="100vh"
        bgGradient="linear-gradient(to bottom, purple.500 50%, white 50%)"
        pos="absolute"
      >
        <Box
          bg="white"
          borderRadius="lg"
          boxShadow="md"
          p={4}
          maxW="90%"
          mx="auto"
          my="80px"
          //   h="100%"
          //   mt={-8}
        >
          <NavigationBar />
          <Grid
            h="100vh"
            templateRows="repeat(6, 1fr)"
            templateColumns="repeat(6, 1fr)"
            gap={4}
          >
            <GridItem
              colSpan={2}
              h="200"
              borderRadius="lg"
              boxShadow="xl" // Add box shadow for depth
              transform="rotateX(10deg)" // Tilt the card
              transition="transform 0.3s ease-in-out" // Add transition for smooth effect
              _hover={{
                transform: "rotateX(0deg)", // Reset transform on hover
              }}
            >
              <Box
                h="50%"
                bgGradient={[
                  // 'linear(to-tr, teal.300, yellow.400)',
                  "linear(to-t, gray.50 50%, teal.300 50%)",
                  "linear(to-b, orange.100, purple.300)",
                ]}
                display={"flex"}
                alignItems={"flex-end"}
                //   justifyContent={'center'}
              >
                <PiAddressBookFill
                  style={{ marginRight: "7px", marginLeft: "60px" }}
                  size={40}
                />
                <Text color="" fontSize="lg" fontWeight="bold">
                  Total Resources
                </Text>
              </Box>
              <Box
                display={"flex"}
                justifyContent={"space-between"}
                alignItems={"center"}
                p={7}
              >
                <Heading>{totalDataCount}</Heading>
                <BsGraphUpArrow size={25} />
              </Box>
            </GridItem>
            <GridItem
              colSpan={2}
              h="200"
              borderRadius="lg"
              boxShadow="xl" // Add box shadow for depth
              transform="rotateX(10deg)" // Tilt the card
              transition="transform 0.3s ease-in-out" // Add transition for smooth effect
              _hover={{
                transform: "rotateX(0deg)", // Reset transform on hover
              }}
            >
              <Box
                h="50%"
                bgGradient={[
                  // 'linear(to-tr, teal.300, yellow.400)',
                  "linear(to-t, gray.50 50%, teal.300 50%)",
                  "linear(to-b, orange.100, purple.300)",
                ]}
                display={"flex"}
                alignItems={"flex-end"}
                //   justifyContent={'center'}
              >
                <FaNoteSticky
                  style={{ marginRight: "7px", marginLeft: "60px" }}
                  size={40}
                />
                <Text color="" fontSize="lg" fontWeight="bold">
                  Total Test Completed
                </Text>
              </Box>
              <Box
                display={"flex"}
                justifyContent={"space-between"}
                alignItems={"center"}
                p={7}
              >
                <Heading>6</Heading>
                <BsGraphUpArrow size={25} />
              </Box>
            </GridItem>
            <GridItem
              colSpan={2}
              h="200"
              borderRadius="lg"
              boxShadow="xl" // Add box shadow for depth
              transform="rotateX(10deg)" // Tilt the card
              transition="transform 0.3s ease-in-out" // Add transition for smooth effect
              _hover={{
                transform: "rotateX(0deg)", // Reset transform on hover
              }}
            >
              <Box
                h="50%"
                bgGradient={[
                  // 'linear(to-tr, teal.300, yellow.400)',
                  "linear(to-t, gray.50 50%, teal.300 50%)",
                  "linear(to-b, orange.100, purple.300)",
                ]}
                display={"flex"}
                alignItems={"flex-end"}
                //   justifyContent={'center'}
              >
                <FaNoteSticky
                  style={{ marginRight: "7px", marginLeft: "60px" }}
                  size={40}
                />
                <Text color="" fontSize="lg" fontWeight="bold">
                  Total Open Test
                </Text>
              </Box>
              <Box
                display={"flex"}
                justifyContent={"space-between"}
                alignItems={"center"}
                p={7}
              >
                <Heading>{totalAllocatedCount}</Heading>
                <BsGraphUpArrow size={25} />
              </Box>
            </GridItem>
            <GridItem
              colSpan={4}
              rowSpan={12}
              bg="white.200"
              borderRadius="lg"
              boxShadow="xl" // Add box shadow for depth
              transform="rotateX(10deg)" // Tilt the card
              transition="transform 0.3s ease-in-out" // Add transition for smooth effect
              _hover={{
                transform: "rotateX(0deg)", // Reset transform on hover
              }}
            >
              <TableContainer style={{ maxHeight: "450px", overflowY: "auto" }}>
                <Table variant="simple">
                  <TableHeader
                    selectedStatus={selectedStatus}
                    setSelectedStatus={setSelectedStatus}
                  />
                  {/* <Box maxH="400" overflowY="auto"> */}
                  <Tbody>
                    {filteredData.map((item: any, index) => (
                      <Tr
                        key={item.id}
                        _hover={{ background: "gray.100", cursor: "pointer" }}
                        onClick={() => console.log("data")} // Add your click functionality here
                      >
                        <Td fontSize="1xl"
            fontWeight="bold">{item.status === "Allocated" ? item.EQL : "Not Generated"}</Td>
                        <Td>{item.name}</Td>
                        <Td>{item.skills}</Td>
                        
                        <Td>
                          <Badge variant="solid" colorScheme={item.status == "Available" ? "green" : "red"}>
                            {item.status  == "Leave" ? "Not Available" : item.status}
                          </Badge>
                        </Td>
                        <Td><SummaryModel data={item} /></Td>
                      </Tr>
                    ))}
                  </Tbody>
                  {/* </Box> */}
                </Table>
              </TableContainer>
            </GridItem>
            <GridItem
              colSpan={2}
              rowSpan={12}
              bg="gray.200"
              borderRadius="lg"
              boxShadow="xl" // Add box shadow for depth
              transform="rotateX(10deg)" // Tilt the card
              transition="transform 0.3s ease-in-out" // Add transition for smooth effect
              _hover={{
                transform: "rotateX(0deg)", // Reset transform on hover
              }}
            >
              <Box
                h={"15%"}
                bg={"purple.500"}
                display="flex"
                alignItems={"center"}
                justifyContent={"center"}
              >
                <Text color={"white"} fontSize={20} fontWeight="bold">
                  Summary Report
                </Text>
              </Box>
              <Box>
                <Chart
                  chartType="PieChart"
                  data={data}
                  options={options}
                  width={"100%"}
                  height={"360px"}
                />
              </Box>
              <Box bg={"white"} display={"flex"} alignItems={"center"}>
                <Text>Total Available:</Text>
              <Button
              bg={"blue"}
              colorScheme="blue"
              // variant="ghost"
              // onClick={""}
              color={"white"}
              marginBottom={"10px"}
              marginLeft={"7px"}
              onClick={handleDownloadCSV}
            >
              Download CSV
            </Button>
              </Box>
              
            </GridItem>
          </Grid>
        </Box>
      </Box>
    </>
  );
}

export default Summary;
